import { AadHttpClient, IHttpClientOptions } from '@microsoft/sp-http';
import { IEmbedConfig, IPowerBiService, IReportSummary, PowerBiErrorKind, PowerBiServiceError } from '../models/IPowerBiModels';
import { API_TIMEOUT_MS, API_ROUTES } from './apiConfig';

const GUID_PATTERN: RegExp = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
// Date ISO 8601 avec fuseau obligatoire : sans fuseau, le navigateur interpréterait l'heure en heure locale.
const ISO_DATE_WITH_TZ_PATTERN: RegExp = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:\d{2})$/;

/**
 * Valide et normalise l'URL de base saisie dans le panneau de propriétés.
 * Retourne undefined si elle est vide ou invalide (HTTPS obligatoire, http://localhost toléré pour les tests).
 */
export function normalizeApiBaseUrl(value: string | undefined): string | undefined {
  const trimmed = (value || '').trim().replace(/\/+$/, '');
  if (!trimmed) {
    return undefined;
  }
  try {
    const url = new URL(trimmed);
    const isLocalhost = url.protocol === 'http:' && url.hostname === 'localhost';
    if ((url.protocol !== 'https:' && !isLocalhost) || url.search || url.hash) {
      return undefined;
    }
    return trimmed;
  } catch {
    return undefined;
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === 'string' && value.trim().length > 0;
}

function invalidResponse(detail: string): PowerBiServiceError {
  return new PowerBiServiceError('invalidResponse', `Réponse de l'API invalide : ${detail}`);
}

/** Valide la réponse de POST /api/embed. Exporté pour que la simulation passe par la même validation. */
export function parseEmbedConfig(json: unknown): IEmbedConfig {
  if (!isRecord(json)) {
    throw invalidResponse('objet JSON attendu');
  }
  const { reportId, embedUrl, embedToken, expiration } = json;

  if (!isNonEmptyString(reportId) || !GUID_PATTERN.test(reportId)) {
    throw invalidResponse('reportId doit être le GUID du rapport Power BI');
  }
  if (!isNonEmptyString(embedUrl)) {
    throw invalidResponse('embedUrl manquant');
  }
  let embedHost: string;
  try {
    const url = new URL(embedUrl);
    embedHost = url.protocol === 'https:' ? url.hostname.toLowerCase() : '';
  } catch {
    embedHost = '';
  }
  if (!/(^|\.)powerbi\.com$/.test(embedHost)) {
    throw invalidResponse('embedUrl doit être une URL https du domaine powerbi.com');
  }
  if (!isNonEmptyString(embedToken)) {
    throw invalidResponse('embedToken manquant');
  }
  if (!isNonEmptyString(expiration) || !ISO_DATE_WITH_TZ_PATTERN.test(expiration) || isNaN(Date.parse(expiration))) {
    throw invalidResponse('expiration doit être une date ISO 8601 avec fuseau (ex. 2026-12-01T15:00:00Z)');
  }

  return { reportId, embedUrl, embedToken, expiration: new Date(expiration) };
}

/** Valide la réponse de GET /api/reports. */
export function parseReportList(json: unknown): IReportSummary[] {
  if (!Array.isArray(json)) {
    throw invalidResponse('tableau de rapports attendu');
  }
  return json.map((item: unknown) => {
    if (!isRecord(item) || !isNonEmptyString(item.id) || !isNonEmptyString(item.name)) {
      throw invalidResponse('chaque rapport doit avoir un id et un name');
    }
    return { id: item.id, name: item.name };
  });
}

/** Convertit une réponse HTTP en erreur, à partir du statut et du corps JSON { error: { code, message, correlationId } }. */
export function errorFromHttpStatus(status: number, body: unknown): PowerBiServiceError {
  let kind: PowerBiErrorKind;
  if (status === 400) {
    kind = 'badRequest';
  } else if (status === 401) {
    kind = 'unauthorized';
  } else if (status === 403) {
    kind = 'forbidden';
  } else if (status === 404) {
    kind = 'notFound';
  } else if (status >= 400 && status < 500 && status !== 408 && status !== 429) {
    kind = 'badRequest';
  } else {
    kind = 'server';
  }

  const error = isRecord(body) && isRecord(body.error) ? body.error : {};
  return new PowerBiServiceError(kind, `L'API a répondu HTTP ${status}`, {
    status,
    code: isNonEmptyString(error.code) ? error.code : undefined,
    correlationId: isNonEmptyString(error.correlationId) ? error.correlationId : undefined
  });
}

/** Réponse HTTP minimale attendue du transport (HttpClientResponse SPFx ou Response du navigateur). */
export interface IApiResponse {
  ok: boolean;
  status: number;
  json(): Promise<unknown>;
}

/** Transport HTTP : AadHttpClient en production, fetch vers le simulateur local en développement. */
export interface IApiTransport {
  fetch(url: string, options: IHttpClientOptions): Promise<IApiResponse>;
}

/**
 * Transport de production. L'AadHttpClient fourni par SPFx obtient silencieusement, pour l'utilisateur déjà
 * connecté à SharePoint, un jeton Entra ID délégué dont l'audience est la ressource de l'API (API_RESOURCE),
 * et l'ajoute en en-tête Authorization: Bearer. Aucun jeton n'est manipulé ni stocké par ce code.
 */
export function createAadTransport(client: AadHttpClient): IApiTransport {
  return { fetch: (url, options) => client.fetch(url, AadHttpClient.configurations.v1, options) };
}

/** Client de l'API d'intégration Power BI (Azure Function). Identique en production et en simulation. */
export class PowerBiApiService implements IPowerBiService {
  public constructor(private readonly _transport: IApiTransport, private readonly _baseUrl: string) {}

  public async getReports(): Promise<IReportSummary[]> {
    const json = await this._request('GET', API_ROUTES.reports);
    return parseReportList(json);
  }

  public async getEmbedConfig(reportId: string): Promise<IEmbedConfig> {
    const json = await this._request('POST', API_ROUTES.embed, { reportId });
    return parseEmbedConfig(json);
  }

  private async _request(method: 'GET' | 'POST', route: string, body?: unknown): Promise<unknown> {
    const controller = new AbortController();
    const timeout = window.setTimeout(() => controller.abort(), API_TIMEOUT_MS);
    const options: IHttpClientOptions = {
      method,
      headers: body === undefined ? { Accept: 'application/json' } : { Accept: 'application/json', 'Content-Type': 'application/json' },
      body: body === undefined ? undefined : JSON.stringify(body),
      signal: controller.signal
    };

    let response: IApiResponse;
    try {
      response = await this._transport.fetch(`${this._baseUrl}${route}`, options);
    } catch (error) {
      throw this._toTransportError(error);
    } finally {
      window.clearTimeout(timeout);
    }

    let json: unknown;
    try {
      json = await response.json();
    } catch {
      json = undefined;
    }

    if (!response.ok) {
      throw errorFromHttpStatus(response.status, json);
    }
    if (json === undefined) {
      throw invalidResponse('corps JSON absent ou illisible');
    }
    return json;
  }

  /** Erreur survenue avant toute réponse HTTP : obtention du jeton Entra ID, CORS, réseau ou délai dépassé. */
  private _toTransportError(error: unknown): PowerBiServiceError {
    const message = error instanceof Error ? error.message : '';
    // Les échecs d'obtention du jeton (consentement absent, ressource inconnue…) remontent avec un code AADSTS.
    if (/AADSTS|consent|token/i.test(message)) {
      return new PowerBiServiceError('unauthorized', 'Impossible d\'obtenir un jeton Entra ID pour l\'API', {
        code: (/AADSTS\d+/.exec(message) || [])[0]
      });
    }
    return new PowerBiServiceError('network', 'API injoignable (réseau, CORS ou délai dépassé)');
  }
}
