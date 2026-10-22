import * as React from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { factories, models, service as pbi, IReportEmbedConfiguration, Page, Report } from 'powerbi-client';
import styles from './PowerBiViewer.module.scss';
import strings from '../loc/strings';
import { IEmbedConfig, IPowerBiService, PowerBiServiceError } from '../models/IPowerBiModels';

export interface IPowerBiViewerProps {
  /** undefined tant que l'URL du service n'est pas configurée. */
  service: IPowerBiService | undefined;
  reportId: string;
  /** Hauteur fixe en pixels ; undefined = hauteur automatique. */
  fixedHeight?: number;
  isEditMode: boolean;
  /** Renseignée uniquement en mode simulation (builds de développement) : affichée dans un bandeau. */
  simulationUrl?: string;
}

type ViewState = { status: 'loading' } | { status: 'ready' } | { status: 'error'; error: PowerBiServiceError };

const RENEWAL_MARGIN_MS: number = 5 * 60 * 1000;
const MIN_RENEWAL_DELAY_MS: number = 10 * 1000;
const RETRY_BASE_DELAY_MS: number = 15 * 1000;
const MAX_TIMER_DELAY_MS: number = 0x7fffffff;
const DEFAULT_PAGE_RATIO: number = 9 / 16;
const MIN_HEIGHT: number = 300;

// Instance dédiée du SDK (plutôt que window.powerbi) : c'est elle qui crée et pilote l'iframe du rapport.
const powerBiService: pbi.Service = new pbi.Service(factories.hpmFactory, factories.wpmpFactory, factories.routerFactory);

/** Délai avant renouvellement : 5 minutes avant l'expiration annoncée par l'API. */
export function computeRenewalDelay(expiration: Date, now: number = Date.now()): number {
  const remaining = expiration.getTime() - now;
  const delay = remaining - RENEWAL_MARGIN_MS;
  // Jeton dont la durée de vie est inférieure à la marge : renouvellement à mi-parcours, sans boucle serrée.
  return Math.min(delay >= MIN_RENEWAL_DELAY_MS ? delay : Math.max(remaining / 2, MIN_RENEWAL_DELAY_MS), MAX_TIMER_DELAY_MS);
}

function toServiceError(error: unknown): PowerBiServiceError {
  return error instanceof PowerBiServiceError
    ? error
    : new PowerBiServiceError('embed', error instanceof Error ? error.message : 'Erreur inattendue');
}

/** Hauteur visible de la zone de défilement de la page SharePoint (ou de la fenêtre à défaut). */
function getVisibleHeight(element: HTMLElement): number {
  for (let node = element.parentElement; node; node = node.parentElement) {
    const overflowY = window.getComputedStyle(node).overflowY;
    if ((overflowY === 'auto' || overflowY === 'scroll') && node.clientHeight > 0) {
      return node.clientHeight;
    }
  }
  return window.innerHeight;
}

const PowerBiViewer: React.FC<IPowerBiViewerProps> = ({ service, reportId, fixedHeight, isEditMode, simulationUrl }) => {
  const wrapperRef = React.useRef<HTMLDivElement>(null);
  const containerRef = React.useRef<HTMLDivElement>(null);
  const reportRef = React.useRef<Report | undefined>(undefined);

  const [view, setView] = React.useState<ViewState>({ status: 'loading' });
  const [pageRatio, setPageRatio] = React.useState<number>(DEFAULT_PAGE_RATIO);
  const [box, setBox] = React.useState({ width: 0, visibleHeight: 0 });

  // Surveille la largeur de la section et la hauteur visible de la page (redimensionnement, changement de mise en page).
  React.useLayoutEffect(() => {
    const wrapper = wrapperRef.current;
    if (!wrapper) {
      return undefined;
    }
    const measure = (): void => {
      const width = wrapper.clientWidth;
      const visibleHeight = getVisibleHeight(wrapper);
      setBox(previous => (previous.width === width && previous.visibleHeight === visibleHeight ? previous : { width, visibleHeight }));
    };
    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(wrapper);
    window.addEventListener('resize', measure);
    return () => {
      observer.disconnect();
      window.removeEventListener('resize', measure);
    };
  }, []);

  /*
   * Dimensionnement :
   * - hauteur automatique : hauteur idéale = largeur × proportions de la page du rapport,
   *   plafonnée à la hauteur visible de l'écran ;
   * - FitToWidth lorsque la page tient dans la hauteur disponible (aucun défilement),
   *   FitToPage lorsqu'elle la dépasse légèrement (page entière visible, pas de double défilement),
   *   FitToWidth à nouveau pour les pages très hautes, qui deviendraient illisibles une fois réduites.
   */
  const idealHeight = box.width * pageRatio;
  const maxHeight = fixedHeight ?? Math.max(box.visibleHeight, MIN_HEIGHT);
  const frameHeight = fixedHeight ?? Math.round(Math.min(Math.max(idealHeight, MIN_HEIGHT), maxHeight));
  const displayOption = idealHeight <= maxHeight || idealHeight > 2 * maxHeight
    ? models.DisplayOption.FitToWidth
    : models.DisplayOption.FitToPage;
  const displayOptionRef = React.useRef(displayOption);

  React.useEffect(() => {
    displayOptionRef.current = displayOption;
    const report = reportRef.current;
    if (report) {
      report.updateSettings({ layoutType: models.LayoutType.Custom, customLayout: { displayOption } })
        .catch(() => undefined); // Rapport pas encore chargé : l'option sera réappliquée après "loaded".
    }
  }, [displayOption]);

  // Cycle de vie complet d'un rapport : chargement, intégration, renouvellement du jeton, nettoyage.
  React.useEffect(() => {
    const container = containerRef.current;
    if (!service || !reportId || !container) {
      return undefined;
    }

    let closed = false;
    let current: IEmbedConfig | undefined;
    let report: Report | undefined;
    let renewalTimer: number | undefined;
    let pendingRenewal: Promise<void> | undefined;
    let failedRenewals = 0;

    setView({ status: 'loading' });
    setPageRatio(DEFAULT_PAGE_RATIO);

    const teardown = (): void => {
      closed = true;
      window.clearTimeout(renewalTimer);
      if (report) {
        powerBiService.reset(container);
        report = undefined;
        reportRef.current = undefined;
      }
    };

    const fail = (error: PowerBiServiceError): void => {
      if (closed) {
        return;
      }
      teardown();
      setView({ status: 'error', error });
    };

    const applyPageSize = (page: Page | undefined): void => {
      const size = page?.defaultSize;
      if (size?.width && size.height) {
        setPageRatio(size.height / size.width);
      }
    };

    function scheduleRenewal(delay: number): void {
      window.clearTimeout(renewalTimer);
      renewalTimer = window.setTimeout(() => {
        renewToken().catch(() => undefined);
      }, delay);
    }

    function onRenewed(next: IEmbedConfig): void {
      current = next;
      failedRenewals = 0;
      scheduleRenewal(computeRenewalDelay(next.expiration));
    }

    function onRenewalFailed(error: PowerBiServiceError, previous: IEmbedConfig): void {
      const remaining = previous.expiration.getTime() - Date.now();
      if (error.kind !== 'network' && error.kind !== 'server') {
        fail(error);
      } else if (remaining <= 0) {
        fail(new PowerBiServiceError('tokenExpired', 'Jeton expiré, renouvellement impossible', {
          code: error.code,
          correlationId: error.correlationId
        }));
      } else {
        failedRenewals++;
        scheduleRenewal(Math.min(RETRY_BASE_DELAY_MS * Math.pow(2, failedRenewals - 1), remaining));
      }
    }

    /*
     * Renouvellement de l'Embed Token : nouvel appel à l'API d'intégration (qui revérifie les droits de l'utilisateur),
     * puis setAccessToken() sur le rapport existant, sans rechargement. Un rechargement n'a lieu que si le
     * jeton précédent avait déjà expiré (onglet en veille, poste suspendu…).
     * Erreurs réseau / 5xx : nouvelles tentatives espacées tant que le jeton courant reste valide.
     * Erreurs 400 / 401 / 403 / 404 / réponse invalide : droits retirés ou configuration invalide, le rapport est fermé.
     */
    async function doRenew(): Promise<void> {
      const previous = current;
      if (closed || !previous || !service) {
        return;
      }
      const wasExpired = previous.expiration.getTime() <= Date.now();
      try {
        const next = await service.getEmbedConfig(reportId);
        if (closed) {
          return;
        }
        if (next.reportId !== previous.reportId) {
          throw new PowerBiServiceError('invalidResponse', 'reportId différent lors du renouvellement');
        }
        // Un jeton déjà expiré provoquerait une boucle TokenExpired → renouvellement → rechargement.
        if (next.expiration.getTime() <= Date.now()) {
          throw new PowerBiServiceError('tokenExpired', 'Jeton renvoyé déjà expiré');
        }
        if (report) {
          await report.setAccessToken(next.embedToken);
          if (wasExpired) {
            await report.reload();
          }
        }
        if (!closed) {
          onRenewed(next);
        }
      } catch (e) {
        if (!closed) {
          onRenewalFailed(toServiceError(e), previous);
        }
      }
    }

    // Un seul renouvellement à la fois : minuterie, retour sur l'onglet et erreur TokenExpired peuvent coïncider.
    function renewToken(): Promise<void> {
      if (!pendingRenewal) {
        pendingRenewal = doRenew().then(() => {
          pendingRenewal = undefined;
        });
      }
      return pendingRenewal;
    }

    // Les minuteries sont suspendues ou retardées en arrière-plan : contrôle au retour sur l'onglet.
    const onVisibilityChange = (): void => {
      if (document.visibilityState === 'visible' && current && current.expiration.getTime() - Date.now() < RENEWAL_MARGIN_MS) {
        renewToken().catch(() => undefined);
      }
    };

    const embedReport = (config: IEmbedConfig): Report => {
      const embedConfiguration: IReportEmbedConfiguration = {
        type: 'report',
        tokenType: models.TokenType.Embed,
        id: config.reportId,
        embedUrl: config.embedUrl,
        accessToken: config.embedToken,
        permissions: models.Permissions.Read,
        settings: {
          background: models.BackgroundType.Transparent,
          layoutType: models.LayoutType.Custom,
          customLayout: { displayOption: displayOptionRef.current }
        }
      };
      const embedded = powerBiService.embed(container, embedConfiguration) as Report;
      let rendered = false;

      embedded.on('loaded', () => {
        embedded.getActivePage().then(applyPageSize).catch(() => undefined);
        embedded.updateSettings({ layoutType: models.LayoutType.Custom, customLayout: { displayOption: displayOptionRef.current } })
          .catch(() => undefined);
      });
      embedded.on('pageChanged', event => {
        applyPageSize((event.detail as { newPage?: Page } | undefined)?.newPage);
      });
      embedded.on('rendered', () => {
        rendered = true;
        setView(previous => (previous.status === 'ready' ? previous : { status: 'ready' }));
      });
      embedded.on('error', event => {
        const detail = (event.detail || {}) as models.IError;
        // Journal de diagnostic sans aucun jeton : message, code et identifiant de requête Power BI uniquement.
        console.warn('[PowerBiViewer] Erreur Power BI', {
          message: detail.message,
          errorCode: detail.errorCode,
          level: detail.level,
          requestId: detail.technicalDetails?.requestId
        });
        if (/TokenExpired/i.test(`${detail.message} ${detail.detailedMessage || ''}`)) {
          renewToken().catch(() => undefined);
        } else if (detail.level === models.TraceType.Fatal || (detail.level === undefined && !rendered)) {
          fail(new PowerBiServiceError('embed', detail.message || 'Erreur Power BI', {
            code: detail.errorCode,
            correlationId: detail.technicalDetails?.requestId
          }));
        }
      });
      return embedded;
    };

    const load = async (): Promise<void> => {
      try {
        const config = await service.getEmbedConfig(reportId);
        if (closed) {
          return;
        }
        current = config;
        report = embedReport(config);
        reportRef.current = report;
        scheduleRenewal(computeRenewalDelay(config.expiration));
      } catch (e) {
        fail(toServiceError(e));
      }
    };

    document.addEventListener('visibilitychange', onVisibilityChange);
    load().catch(() => undefined);

    return () => {
      document.removeEventListener('visibilitychange', onVisibilityChange);
      teardown();
    };
  }, [service, reportId]);

  const isConfigured = !!service && !!reportId;
  const error = view.status === 'error' ? view.error : undefined;
  const supportInfo = error && [
    error.status ? `HTTP ${error.status}` : '',
    error.code ? `code ${error.code}` : '',
    error.correlationId ? `référence ${error.correlationId}` : ''
  ].filter(Boolean).join(' · ');

  return (
    <div ref={wrapperRef} className={styles.viewer}>
      {simulationUrl && (
        <MessageBar messageBarType={MessageBarType.warning}>
          {strings.viewer.simulationNotice} ({simulationUrl})
        </MessageBar>
      )}

      {!isConfigured && (
        <MessageBar messageBarType={MessageBarType.info}>
          {isEditMode ? strings.viewer.notConfiguredEdit : strings.viewer.notConfiguredRead}
        </MessageBar>
      )}

      {isConfigured && error && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          {strings.errors[error.kind]}
          {supportInfo && <div className={styles.supportInfo}>{strings.viewer.errorDetails} : {supportInfo}</div>}
        </MessageBar>
      )}

      <div className={`${styles.frame} ${!isConfigured || error ? styles.hidden : ''}`} style={{ height: frameHeight }}>
        <div ref={containerRef} className={`${styles.reportContainer}`} />

        {view.status === 'loading' && (
          <div className={styles.loading}>
            <Spinner size={SpinnerSize.medium} ariaLabel={strings.viewer.loading} />
          </div>
        )}
      </div>
    </div>
  );
};

export default PowerBiViewer;
