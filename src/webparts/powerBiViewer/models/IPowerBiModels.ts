/** Entrée du catalogue renvoyé par GET /api/reports (contrat provisoire, cf. docs/INTEGRATION_API.md). */
export interface IReportSummary {
  /** Identifiant à renvoyer tel quel dans POST /api/embed. */
  id: string;
  name: string;
}

/** Informations d'intégration renvoyées par POST /api/embed, après validation. */
export interface IEmbedConfig {
  /** GUID du rapport Power BI (utilisé par powerbi-client). */
  reportId: string;
  embedUrl: string;
  embedToken: string;
  expiration: Date;
}

export type PowerBiErrorKind =
  | 'badRequest'
  | 'unauthorized'
  | 'forbidden'
  | 'notFound'
  | 'server'
  | 'network'
  | 'invalidResponse'
  | 'embed'
  | 'tokenExpired';

/** Erreur normalisée. Ne contient jamais de jeton : seuls le code et la référence de corrélation sont conservés. */
export class PowerBiServiceError extends Error {
  public readonly kind: PowerBiErrorKind;
  public readonly status?: number;
  public readonly code?: string;
  public readonly correlationId?: string;

  public constructor(kind: PowerBiErrorKind, message: string, details?: { status?: number; code?: string; correlationId?: string }) {
    super(message);
    // Nécessaire avec une cible ES5 pour que instanceof fonctionne sur une sous-classe d'Error.
    Object.setPrototypeOf(this, PowerBiServiceError.prototype);
    this.name = 'PowerBiServiceError';
    this.kind = kind;
    this.status = details?.status;
    this.code = details?.code;
    this.correlationId = details?.correlationId;
  }
}

/** Contrat commun au service réel (API d'intégration) et au service de simulation. */
export interface IPowerBiService {
  getReports(): Promise<IReportSummary[]>;
  /** Utilisé pour le chargement initial et pour chaque renouvellement de l'Embed Token. */
  getEmbedConfig(reportId: string): Promise<IEmbedConfig>;
}
