import { IHttpClientOptions } from '@microsoft/sp-http';
import { IApiTransport } from './PowerBiApiService';

/**
 * Transport de SIMULATION, pour le développement sans l'API d'intégration réelle.
 *
 * Il n'est chargé que dans les builds de développement (cf. PowerBiViewerWebPart, garde `DEBUG`) : il est
 * absent du package de production. Il adresse les mêmes requêtes que la production au simulateur local
 * (tools/Start-BackendSimulator.ps1), qui reproduit le contrat de l'API et génère de vrais Embed Tokens.
 *
 * Seule différence avec la production : aucun jeton Entra ID n'est disponible pour le simulateur. Un jeton
 * factice est envoyé pour vérifier que le backend exige bien l'en-tête Authorization.
 */
export function createSimulationTransport(): IApiTransport {
  return {
    fetch: (url: string, options: IHttpClientOptions) => {
      const headers = new Headers(options.headers as HeadersInit);
      headers.set('Authorization', 'Bearer simulation-locale');
      return window.fetch(url, { ...options, headers } as RequestInit);
    }
  };
}
