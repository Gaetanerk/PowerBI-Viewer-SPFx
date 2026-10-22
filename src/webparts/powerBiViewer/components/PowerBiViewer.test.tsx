// Cycle de vie de la visionneuse : powerbi-client est remplacé par une doublure, les minuteries sont simulées.

interface IFakeReport {
  config: { id: string; accessToken: string; tokenType: number; permissions: number };
  handlers: Record<string, (event: { detail?: unknown }) => void>;
  on(name: string, handler: (event: { detail?: unknown }) => void): void;
  setAccessToken: jest.Mock;
  reload: jest.Mock;
  updateSettings: jest.Mock;
  getActivePage: jest.Mock;
}

const mockPowerBi = { embedded: [] as IFakeReport[], resets: 0 };

jest.mock('@microsoft/sp-http', () => ({ AadHttpClient: { configurations: { v1: {} } } }));
jest.mock('powerbi-client', () => ({
  models: {
    TokenType: { Aad: 0, Embed: 1 },
    Permissions: { Read: 0 },
    BackgroundType: { Transparent: 1 },
    LayoutType: { Custom: 2 },
    DisplayOption: { FitToPage: 0, FitToWidth: 1 },
    TraceType: { Fatal: 6 }
  },
  factories: {},
  service: {
    Service: class {
      public embed(_element: HTMLElement, config: IFakeReport['config']): IFakeReport {
        const report: IFakeReport = {
          config,
          handlers: {},
          on(name, handler) { this.handlers[name] = handler; },
          setAccessToken: jest.fn(() => Promise.resolve()),
          reload: jest.fn(() => Promise.resolve()),
          updateSettings: jest.fn(() => Promise.resolve()),
          getActivePage: jest.fn(() => Promise.resolve(undefined))
        };
        mockPowerBi.embedded.push(report);
        return report;
      }
      public reset(): void { mockPowerBi.resets++; }
    }
  }
}));

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act } from 'react-dom/test-utils';
import PowerBiViewer, { computeRenewalDelay } from './PowerBiViewer';
import strings from '../loc/strings';
import { IEmbedConfig, IPowerBiService, PowerBiServiceError } from '../models/IPowerBiModels';
import { errorFromHttpStatus } from '../services/PowerBiApiService';

const MINUTE = 60 * 1000;
const REPORT_ID = 'f6bfd646-b718-44dc-a378-b73e6b528204';

describe('computeRenewalDelay', () => {
  const now = Date.UTC(2026, 11, 1, 14, 0, 0);

  it("renouvelle 5 minutes avant l'expiration", () => {
    expect(computeRenewalDelay(new Date(now + 60 * MINUTE), now)).toBe(55 * MINUTE);
  });

  it('renouvelle à mi-parcours un jeton de courte durée', () => {
    expect(computeRenewalDelay(new Date(now + 4 * MINUTE), now)).toBe(2 * MINUTE);
  });

  it('ne descend jamais sous 10 secondes, même pour un jeton déjà expiré', () => {
    expect(computeRenewalDelay(new Date(now - MINUTE), now)).toBe(10000);
  });
});

describe('PowerBiViewer', () => {
  let host: HTMLDivElement;

  beforeAll(() => {
    (window as unknown as { ResizeObserver: unknown }).ResizeObserver = class {
      public observe(): void { /* jsdom : pas de mise en page */ }
      public disconnect(): void { /* idem */ }
    };
  });

  beforeEach(() => {
    jest.useFakeTimers();
    mockPowerBi.embedded = [];
    mockPowerBi.resets = 0;
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    // act() force l'exécution immédiate des nettoyages d'effets, différés par React 17.
    act(() => {
      ReactDom.unmountComponentAtNode(host);
    });
    host.remove();
    jest.useRealTimers();
  });

  async function advance(ms: number): Promise<void> {
    await act(async () => {
      jest.advanceTimersByTime(ms);
    });
    // Laisse se résoudre les promesses enchaînées après la minuterie.
    await act(async () => {
      await Promise.resolve();
      await Promise.resolve();
    });
    // MessageBar (Fluent UI) affiche son texte après un setTimeout(0) (DelayedRender).
    act(() => {
      jest.advanceTimersByTime(1);
    });
  }

  function renderViewer(service: IPowerBiService | undefined, simulationUrl?: string): void {
    act(() => {
      ReactDom.render(
        <PowerBiViewer service={service} reportId={service ? 'rapport' : ''} isEditMode={false} simulationUrl={simulationUrl} />,
        host
      );
    });
  }

  type FakeResponse = 'ok' | 'expired' | PowerBiServiceError;

  /** Faux service : réponses successives après 600 ms ; jeton valable 5 min 20 s (renouvellement après ~20 s). */
  function fakeService(...responses: FakeResponse[]): IPowerBiService & { calls: number } {
    const service = {
      calls: 0,
      getReports: () => Promise.resolve([]),
      getEmbedConfig: (): Promise<IEmbedConfig> => {
        const response = responses[Math.min(service.calls, responses.length - 1)];
        service.calls++;
        const token = `TOKEN-${service.calls}`;
        return new Promise((resolve, reject) => window.setTimeout(() => (response instanceof PowerBiServiceError
          ? reject(response)
          : resolve({
            reportId: REPORT_ID,
            embedUrl: 'https://app.powerbi.com/reportEmbed',
            embedToken: token,
            expiration: new Date(Date.now() + (response === 'expired' ? -MINUTE : 5 * MINUTE + 20000))
          })), 600));
      }
    };
    return service;
  }

  it("affiche un message si la WebPart n'est pas configurée", async () => {
    renderViewer(undefined);
    await advance(0);
    expect(host.textContent).toContain(strings.viewer.notConfiguredRead);
  });

  it('intègre le rapport avec un Embed Token en lecture seule', async () => {
    renderViewer(fakeService('ok'));
    await advance(700);
    expect(mockPowerBi.embedded).toHaveLength(1);
    expect(mockPowerBi.embedded[0].config).toMatchObject({ id: REPORT_ID, accessToken: 'TOKEN-1', tokenType: 1, permissions: 0 });
  });

  it('renouvelle le jeton 5 minutes avant son expiration, sans recharger le rapport', async () => {
    renderViewer(fakeService('ok'));
    await advance(700);
    const report = mockPowerBi.embedded[0];

    await advance(20000);
    await advance(700);
    expect(report.setAccessToken).toHaveBeenCalledWith('TOKEN-2');
    expect(report.reload).not.toHaveBeenCalled();
    expect(mockPowerBi.embedded).toHaveLength(1);
  });

  it('renouvelle le jeton sur un événement TokenExpired de Power BI', async () => {
    renderViewer(fakeService('ok'));
    await advance(700);
    const report = mockPowerBi.embedded[0];

    act(() => {
      report.handlers.error({ detail: { message: 'TokenExpired' } });
    });
    await advance(700);
    expect(report.setAccessToken).toHaveBeenCalledWith('TOKEN-2');
  });

  it('ferme le rapport si les droits sont retirés au renouvellement', async () => {
    renderViewer(fakeService('ok', errorFromHttpStatus(403, { error: { code: 'USER_FORBIDDEN', correlationId: 'abc' } })));
    await advance(700);

    await advance(20000);
    await advance(700);
    expect(mockPowerBi.resets).toBe(1);
    expect(host.textContent).toContain(strings.errors.forbidden);
    expect(host.textContent).toContain('HTTP 403 · code USER_FORBIDDEN · référence abc');
  });

  it('signale un jeton renvoyé déjà expiré au lieu de boucler', async () => {
    renderViewer(fakeService('ok', 'expired'));
    await advance(700);

    await advance(20000);
    await advance(700);
    expect(host.textContent).toContain(strings.errors.tokenExpired);
  });

  it.each([
    [400, strings.errors.badRequest],
    [401, strings.errors.unauthorized],
    [404, strings.errors.notFound],
    [503, strings.errors.server]
  ])('HTTP %i au chargement affiche le message adapté', async (status, message) => {
    renderViewer(fakeService(errorFromHttpStatus(status, undefined)));
    await advance(700);
    expect(host.textContent).toContain(message);
    expect(mockPowerBi.embedded).toHaveLength(0);
  });

  it('API injoignable affiche le message adapté', async () => {
    renderViewer(fakeService(new PowerBiServiceError('network', 'injoignable')));
    await advance(700);
    expect(host.textContent).toContain(strings.errors.network);
  });

  it("n'effectue qu'un appel par renouvellement, même au retour sur l'onglet", async () => {
    const service = fakeService('ok');
    renderViewer(service);
    await advance(700);
    expect(service.calls).toBe(1);

    await advance(20000); // la minuterie déclenche le renouvellement (en cours pendant 600 ms)
    act(() => {
      document.dispatchEvent(new Event('visibilitychange'));
    });
    await advance(700);
    expect(service.calls).toBe(2);
  });

  it('signale le mode simulation par un bandeau', async () => {
    renderViewer(fakeService('ok'), 'http://localhost:7071');
    await advance(700);
    expect(host.textContent).toContain(`${strings.viewer.simulationNotice} (http://localhost:7071)`);
  });
});
