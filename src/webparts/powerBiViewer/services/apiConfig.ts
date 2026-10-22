/**
 * Paramètres de raccordement à l'API d'intégration Power BI — PROVISOIRES.
 *
 * Ces valeurs sont fournies par l'équipe qui développe l'API (cf. docs/INTEGRATION_API.md, section J).
 * Elles sont publiques (aucun secret) mais figées dans le package : les modifier impose
 * de régénérer le .sppkg, tout comme la déclaration webApiPermissionRequests associée
 * dans config/package-solution.json. Seule l'URL du service est réglable dans la page.
 */

/**
 * Ressource Entra ID demandée à AadHttpClient : Application ID URI (ex. api://xxxx)
 * ou Application (client) ID de l'enregistrement d'application exposant l'API d'intégration.
 */
export const API_RESOURCE: string = 'api://00000000-0000-0000-0000-000000000000';

/** Routes relatives à l'URL de base configurée dans la WebPart. */
export const API_ROUTES = {
  reports: '/api/reports',
  embed: '/api/embed'
} as const;

/** Délai au-delà duquel un appel à l'API est considéré comme en échec réseau. */
export const API_TIMEOUT_MS: number = 30000;

/** URL du simulateur local (tools/Start-BackendSimulator.ps1), utilisée en mode simulation si aucune URL n'est saisie. */
export const SIMULATION_DEFAULT_URL: string = 'http://localhost:7071';
