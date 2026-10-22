// Vérifie la validation du contrat de l'API décrit dans docs/INTEGRATION_API.md (section D).
jest.mock('@microsoft/sp-http', () => ({ AadHttpClient: { configurations: { v1: {} } } }));

import { IHttpClientOptions } from '@microsoft/sp-http';
import { errorFromHttpStatus, IApiTransport, normalizeApiBaseUrl, parseEmbedConfig, parseReportList, PowerBiApiService } from './PowerBiApiService';
import { PowerBiServiceError } from '../models/IPowerBiModels';

const validResponse = {
  reportId: 'f6bfd646-b718-44dc-a378-b73e6b528204',
  embedUrl: 'https://app.powerbi.com/reportEmbed?reportId=f6bfd646-b718-44dc-a378-b73e6b528204&groupId=be8908da-da25-452e-b220-163f52476cdd',
  embedToken: 'H4sIAAAAAAAEAB2Wxa6E',
  expiration: '2026-12-01T15:00:00Z'
};

function kindOf(action: () => unknown): string | undefined {
  try {
    action();
    return undefined;
  } catch (error) {
    return error instanceof PowerBiServiceError ? error.kind : 'autre';
  }
}

describe('parseEmbedConfig', () => {
  it('accepte une réponse conforme', () => {
    const config = parseEmbedConfig(validResponse);
    expect(config.reportId).toBe(validResponse.reportId);
    expect(config.expiration.toISOString()).toBe('2026-12-01T15:00:00.000Z');
  });

  it('accepte une expiration avec décalage horaire et millisecondes', () => {
    const config = parseEmbedConfig({ ...validResponse, expiration: '2026-12-01T16:00:00.123+01:00' });
    expect(config.expiration.toISOString()).toBe('2026-12-01T15:00:00.123Z');
  });

  it.each([
    ['reportId non GUID', { reportId: 'rapport-001' }],
    ['embedUrl hors powerbi.com', { embedUrl: 'https://exemple.com/reportEmbed' }],
    ['embedUrl en http', { embedUrl: 'http://app.powerbi.com/reportEmbed' }],
    ['embedToken vide', { embedToken: '' }],
    ['expiration sans fuseau', { expiration: '2026-12-01T15:00:00' }],
    ['expiration non ISO', { expiration: '01/12/2026 15:00' }]
  ])('rejette : %s', (_label, override) => {
    expect(kindOf(() => parseEmbedConfig({ ...validResponse, ...override }))).toBe('invalidResponse');
  });
});

describe('parseReportList', () => {
  it('accepte un tableau { id, name }', () => {
    expect(parseReportList([{ id: 'rapport-001', name: 'Suivi', extra: true }])).toEqual([{ id: 'rapport-001', name: 'Suivi' }]);
  });

  it('rejette un objet enveloppe', () => {
    expect(kindOf(() => parseReportList({ value: [] }))).toBe('invalidResponse');
  });
});

describe('errorFromHttpStatus', () => {
  it.each([
    [400, 'badRequest'],
    [401, 'unauthorized'],
    [403, 'forbidden'],
    [404, 'notFound'],
    [429, 'server'],
    [500, 'server'],
    [503, 'server']
  ])('HTTP %i → %s', (status, kind) => {
    expect(errorFromHttpStatus(status, undefined).kind).toBe(kind);
  });

  it('extrait code et correlationId du corps d\'erreur', () => {
    const error = errorFromHttpStatus(403, { error: { code: 'REPORT_FORBIDDEN', message: 'x', correlationId: 'abc-123' } });
    expect(error.code).toBe('REPORT_FORBIDDEN');
    expect(error.correlationId).toBe('abc-123');
  });
});

describe('normalizeApiBaseUrl', () => {
  it.each([
    ['https://func-powerbi.azurewebsites.net/', 'https://func-powerbi.azurewebsites.net'],
    ['  https://api.exemple.fr/powerbi  ', 'https://api.exemple.fr/powerbi'],
    ['http://localhost:7071', 'http://localhost:7071'],
    ['http://func-powerbi.azurewebsites.net', undefined],
    ['https://func-powerbi.azurewebsites.net/?code=secret', undefined],
    ['', undefined]
  ])('%s → %s', (input, expected) => {
    expect(normalizeApiBaseUrl(input)).toBe(expected);
  });
});

describe('PowerBiApiService', () => {
  function transport(status: number, body: unknown): IApiTransport & { calls: Array<{ url: string; options: IHttpClientOptions }> } {
    const calls: Array<{ url: string; options: IHttpClientOptions }> = [];
    return {
      calls,
      fetch: (url, options) => {
        calls.push({ url, options });
        return Promise.resolve({ ok: status < 400, status, json: () => Promise.resolve(body) });
      }
    };
  }

  it('envoie POST {base}/api/embed avec { reportId } en JSON', async () => {
    const fake = transport(200, validResponse);
    const config = await new PowerBiApiService(fake, 'https://func-powerbi.azurewebsites.net').getEmbedConfig('rapport-001');
    expect(config.embedToken).toBe(validResponse.embedToken);
    expect(fake.calls[0].url).toBe('https://func-powerbi.azurewebsites.net/api/embed');
    expect(fake.calls[0].options.method).toBe('POST');
    expect(fake.calls[0].options.body).toBe('{"reportId":"rapport-001"}');
    expect(fake.calls[0].options.headers).toMatchObject({ Accept: 'application/json', 'Content-Type': 'application/json' });
  });

  it('envoie GET {base}/api/reports', async () => {
    const fake = transport(200, [{ id: 'a', name: 'A' }]);
    await new PowerBiApiService(fake, 'http://localhost:7071').getReports();
    expect(fake.calls[0].url).toBe('http://localhost:7071/api/reports');
    expect(fake.calls[0].options.method).toBe('GET');
  });

  it('convertit une réponse 403 en erreur forbidden', async () => {
    const fake = transport(403, { error: { code: 'USER_FORBIDDEN', correlationId: 'x' } });
    await expect(new PowerBiApiService(fake, 'http://localhost:7071').getEmbedConfig('r')).rejects.toMatchObject({ kind: 'forbidden', code: 'USER_FORBIDDEN' });
  });

  it('convertit un échec réseau en erreur network', async () => {
    const failing: IApiTransport = { fetch: () => Promise.reject(new TypeError('Failed to fetch')) };
    await expect(new PowerBiApiService(failing, 'http://localhost:7071').getEmbedConfig('r')).rejects.toMatchObject({ kind: 'network' });
  });
});
