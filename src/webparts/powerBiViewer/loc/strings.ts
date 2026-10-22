import { PowerBiErrorKind } from '../models/IPowerBiModels';

/** Libellés de l'interface, en français (public cible francophone). */
const strings = {
  propertyPane: {
    description: 'Configurez la source du rapport Power BI. Aucun jeton ni secret n\'est enregistré dans la page.',
    connectionGroup: 'Service de rapports',
    displayGroup: 'Affichage',
    mockModeLabel: 'Mode simulation (build de développement uniquement)',
    mockModeOn: 'Activé : simulateur local de l’API',
    mockModeOff: 'Désactivé',
    apiBaseUrlLabel: 'URL de l\'API d\'intégration Power BI',
    apiBaseUrlDescription: 'URL de base, sans /api (ex. https://mon-service.azurewebsites.net).',
    apiBaseUrlSimulationDescription: 'Simulation : URL du simulateur local (vide = http://localhost:7071), lancé par tools/Start-BackendSimulator.ps1.',
    apiBaseUrlInvalid: 'Saisissez une URL HTTPS valide, sans paramètres.',
    reportLabel: 'Rapport Power BI',
    reportLoading: 'Chargement de la liste des rapports…',
    reportManualDescription: 'Identifiant du rapport attendu par l\'API d\'intégration.',
    reportCatalogUnavailable: 'Liste des rapports indisponible : saisissez l\'identifiant fourni par l\'administrateur de l\'API.',
    reportCatalogEmpty: 'Aucun rapport disponible pour votre compte : saisissez l\'identifiant fourni par l\'administrateur de l\'API.',
    reportNotInCatalog: '(absent de la liste)',
    reportRequired: 'L\'identifiant du rapport est obligatoire.',
    autoHeightLabel: 'Hauteur automatique',
    autoHeightOn: 'Adaptée au rapport et à l\'écran',
    autoHeightOff: 'Hauteur fixe',
    customHeightLabel: 'Hauteur (pixels)'
  },
  viewer: {
    notConfiguredEdit: 'La visionneuse n\'est pas configurée. Modifiez la WebPart (icône crayon) pour renseigner l\'URL du service et le rapport à afficher.',
    notConfiguredRead: 'Aucun rapport n\'est encore configuré pour cet emplacement. Contactez le responsable de la page.',
    loading: 'Chargement du rapport…',
    errorDetails: 'Informations pour le support',
    simulationNotice: 'Mode simulation (build de développement) : données fournies par le simulateur local'
  },
  errors: {
    badRequest: 'La demande d\'affichage du rapport est invalide. Vérifiez la configuration de la visionneuse.',
    unauthorized: 'Votre identité n\'a pas pu être vérifiée par le service de rapports. Rechargez la page ; si le problème persiste, contactez l\'administrateur (autorisation d\'accès à l\'API à approuver).',
    forbidden: 'Vous n\'êtes pas autorisé à consulter ce rapport.',
    notFound: 'Le rapport demandé est introuvable ou n\'est plus disponible.',
    server: 'Le service de rapports est momentanément indisponible. Réessayez dans quelques instants en rechargeant la page.',
    network: 'Le service de rapports est inaccessible. Vérifiez votre connexion puis rechargez la page.',
    invalidResponse: 'Le service de rapports a renvoyé une réponse inattendue : le rapport ne peut pas être affiché.',
    embed: 'Le rapport Power BI ne peut pas être affiché.',
    tokenExpired: 'L\'autorisation d\'affichage du rapport a expiré et n\'a pas pu être renouvelée. Rechargez la page.'
  } as Record<PowerBiErrorKind, string>
};

export default strings;
