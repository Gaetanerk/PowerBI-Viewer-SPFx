# Visionneuse Power BI Embedded — WebPart SPFx

WebPart SharePoint Online qui affiche un rapport Power BI Embedded (scénario **« App owns data »**, capacité Fabric, ex. F64) sur les pages de diffusion de l'organisation cliente, y compris pour des utilisateurs externes invités en B2B dans Entra ID.

La WebPart est **uniquement une visionneuse** :

- elle appelle l'API d'intégration (Azure Function) avec `AadHttpClient` (jeton Entra ID délégué, obtenu silencieusement pour l'utilisateur déjà connecté à SharePoint) ;
- l'Azure Function contrôle les droits et renvoie un **Embed Token** en lecture seule ;
- le SDK officiel `powerbi-client` crée l'iframe et affiche le rapport.

Elle ne s'authentifie jamais directement auprès de Power BI, ne génère aucun jeton et ne contient aucun secret.

Le contrat technique avec l'API d'intégration est décrit dans [docs/INTEGRATION_API.md](docs/INTEGRATION_API.md).

## Technologies

| Élément | Version |
|---|---|
| SharePoint Framework (chaîne Heft) | 1.23.2 |
| Node.js | >= 22.14 et < 23 |
| React / React DOM | 17.0.1 |
| TypeScript | 5.8 |
| Fluent UI | `@fluentui/react` 8 |
| SDK Power BI | `powerbi-client` 2.24.1 |
| Appels API | `AadHttpClient` (`@microsoft/sp-http` 1.23.2) |

## Architecture des fichiers

```text
src/webparts/powerBiViewer/
├── PowerBiViewerWebPart.ts            Paramétrage SPFx : panneau de propriétés, catalogue, choix du service
├── PowerBiViewerWebPart.manifest.json
├── components/
│   ├── PowerBiViewer.tsx              Affichage : powerbi-client, renouvellement du jeton, dimensionnement, erreurs
│   ├── PowerBiViewer.module.scss
│   └── PowerBiViewer.test.tsx
├── services/
│   ├── apiConfig.ts                   Ressource Entra ID et routes de l'API d'intégration (PROVISOIRES)
│   ├── PowerBiApiService.ts           Communication avec l'API d'intégration et validation des réponses
│   ├── PowerBiApiService.test.ts
│   └── SimulationTransport.ts         Transport du mode simulation (builds de développement uniquement)
├── models/
│   └── IPowerBiModels.ts              Types partagés et erreur normalisée
└── loc/
    └── strings.ts                     Libellés français

tools/Start-BackendSimulator.ps1       Simulateur local de l'API d'intégration (vrais Embed Tokens)
dev/samples/AdventureWorks Sales.pbix  Rapport d'exemple Microsoft (exclu de Git, retéléchargé par le simulateur)
docs/RECETTE.md                        Procédure de recette du cahier des charges
```

Pour adapter le contrat de l'API, seuls `services/apiConfig.ts` et `services/PowerBiApiService.ts` sont concernés ; les composants d'affichage ne dépendent que de l'interface `IPowerBiService`.

## Installation et commandes

```bash
npm install
npm start            # build de développement + serveur local, à ouvrir dans le workbench hébergé
npx heft test        # compilation, lint et tests unitaires
npm run build        # build de production + package
```

`npm start` sert la WebPart sur `https://localhost:4321` ; ouvrez ensuite `https://<votre-tenant>.sharepoint.com/_layouts/workbench.aspx` (remplacez `{tenantDomain}` dans [config/serve.json](config/serve.json) pour l'ouverture automatique).

Le package est produit dans **`sharepoint/solution/power-bi-viewer.sppkg`**. Pensez à incrémenter `version` dans [config/package-solution.json](config/package-solution.json) à chaque nouvelle livraison.

## Configuration dans SharePoint

Panneau de propriétés de la WebPart :

| Paramètre | Description |
|---|---|
| URL de l'API d'intégration (Azure Function) | URL de base **sans `/api`**, en HTTPS, sans paramètre (ex. `https://func-powerbi.azurewebsites.net`). Une clé de fonction (`?code=`) est refusée : elle serait exposée dans le navigateur. |
| Rapport Power BI | Liste déroulante alimentée par `GET /api/reports` si l'API expose ce catalogue ; sinon champ de saisie de l'identifiant attendu par l'équipe backend. |
| Hauteur automatique | Activée par défaut. Désactivée : hauteur fixe de 300 à 2000 px. |

Aucun jeton, secret ni identifiant de service n'est enregistré dans les propriétés de la page.

### Affichage

- Largeur : 100 % de la section (la WebPart accepte aussi les sections pleine largeur), sans marge ni bordure ajoutée.
- Hauteur automatique : largeur × proportions de la page active du rapport, plafonnée à la hauteur visible de la zone de défilement SharePoint ; recalculée au redimensionnement et au changement de page.
- Ajustement Power BI : `FitToWidth` quand la page tient dans la hauteur disponible, `FitToPage` quand elle la dépasse légèrement (pas de double défilement), `FitToWidth` pour les pages très hautes.
- Pas d'en-tête, de barre d'outils ni de mode plein écran. Filtres et navigation entre pages restent ceux du rapport.

### Mise à jour des rapports

Chaque chargement de page demande de nouvelles informations d'intégration à l'équipe backend, sans cache côté navigateur : la dernière version publiée du rapport s'affiche au prochain chargement, sans redéploiement, tant que l'identifiant du rapport reste valide (rapport republié « en remplacement », pas sous un nouvel identifiant).

### Renouvellement de l'Embed Token

Nouvel appel à `POST /api/embed` environ 5 minutes avant l'expiration annoncée par l'équipe backend, puis `report.setAccessToken()` sans rechargement. Détail du comportement (réessais, perte de droits, onglet en veille) : section G de [docs/INTEGRATION_API.md](docs/INTEGRATION_API.md#g-renouvellement-des-tokens).

## Configuration Entra ID et autorisations

Valeurs à obtenir auprès de l'équipe backend (voir la section J du document d'intégration) :

1. **Ressource Entra ID** : renseigner l'Application ID URI (ou l'Application ID) de l'API dans `API_RESOURCE` ([services/apiConfig.ts](src/webparts/powerBiViewer/services/apiConfig.ts)). La valeur actuelle est un **exemple fictif**.
2. **Permission déléguée** : ajouter dans [config/package-solution.json](config/package-solution.json), sous `solution` :

   ```json
   "webApiPermissionRequests": [
     {
       "resource": "<Nom de l'application API d'intégration dans Entra ID>",
       "scope": "<scope délégué, ex. user_impersonation>"
     }
   ]
   ```

   Cette déclaration n'est volontairement pas présente tant que les valeurs réelles ne sont pas connues : une valeur fictive créerait une demande d'autorisation invalide dans le centre d'administration SharePoint.
3. **Approbation** : après déploiement du package, un administrateur SharePoint approuve la demande dans *Centre d'administration SharePoint › Avancé › Accès à l'API*.

Le rôle `Reports.Read` évoqué par Le backend doit être qualifié : s'il s'agit d'un **scope délégué**, il se déclare dans `webApiPermissionRequests` ; s'il s'agit d'un **rôle d'application** attribué à des utilisateurs ou au groupe autorisé, il ne se déclare pas dans SPFx et apparaît dans la revendication `roles` du jeton.

## Mode simulation

Le mode simulation permet de vérifier la WebPart **de bout en bout, avec un vrai rapport Power BI**, sans l'API d'intégration réelle ni l'environnement du client.

```text
Workbench SharePoint ── WebPart (PowerBiApiService, code de production) ── HTTP ──► Simulateur local (http://localhost:7071)
                                                                                    │  GET /api/reports, POST /api/embed
                                                                                    ▼
                                                                         API REST Power BI (GenerateToken, votre compte)
```

- **Simulateur** : [tools/Start-BackendSimulator.ps1](tools/Start-BackendSimulator.ps1) reproduit le contrat de [docs/INTEGRATION_API.md](docs/INTEGRATION_API.md), comme le fera l'Azure Function :
  - mêmes routes et mêmes formats JSON ;
  - en-tête `Authorization` exigé (401 sinon) ;
  - contrôle du `reportId` côté serveur ;
  - CORS ;
  - erreurs `{ error: { code, message, correlationId } }` ;
  - un **vrai Embed Token en lecture seule généré à chaque appel**, y compris aux renouvellements.

  Il utilise votre compte Power BI à la place du principal de service, et journalise chaque appel dans la console, sans jeton.
- **WebPart** : en mode simulation, elle utilise le même `PowerBiApiService` qu'en production. Seul le transport change : `fetch` vers le simulateur avec un jeton factice, au lieu d'`AadHttpClient` avec un jeton Entra ID.
- **Isolation** : l'interrupteur et le transport de simulation n'existent que dans les builds de développement (`npm start`). Dans le package de production, webpack les retire (garde `DEBUG`).

### Afficher AdventureWorks Sales

Prérequis : un compte Power BI avec licence Pro, PPU ou essai gratuit (création d'un espace de travail). Un fichier `.pbix` ne peut pas être affiché directement depuis le poste : Power BI Embedded n'affiche que des rapports publiés dans le service Power BI. Le simulateur s'en charge.

1. Dans un terminal PowerShell, à la racine du projet :

   ```powershell
   ./tools/Start-BackendSimulator.ps1
   ```

   Au premier lancement, le module `MicrosoftPowerBIMgmt` est installé et une fenêtre de connexion Microsoft s'ouvre. Le simulateur crée ensuite l'espace de travail « Visionneuse Power BI - Tests » et y publie [dev/samples/AdventureWorks Sales.pbix](dev/samples/), le rapport d'exemple Microsoft (données fictives, licence MIT, téléchargé s'il est absent). Laisser la fenêtre ouverte (arrêt : Ctrl+C).
2. Dans un second terminal : `npm start`, puis ouvrir `https://<votre-tenant>.sharepoint.com/_layouts/workbench.aspx`.
3. Ajouter la WebPart, ouvrir ses propriétés (icône crayon), activer **Mode simulation** (URL laissée vide = `http://localhost:7071`) et choisir **AdventureWorks Sales** dans la liste.

Le rapport s'affiche sous un bandeau « Mode simulation » ; chaque appel apparaît dans la console du simulateur.

| Option | Effet |
|---|---|
| `-TokenLifetimeMinutes 10` | Jetons de 10 min : renouvellement observable au bout d'environ 5 min |
| `-SimulateErrors` | Ajoute au catalogue des rapports `[Test]` : accès refusé (403), rapport supprimé (404), service indisponible (503), droits retirés au renouvellement (403) |
| `-PbixPath "C:\chemin\rapport.pbix"` | Publie ce fichier dans l'espace de travail de test (remplacement sous le même identifiant) |
| `-Port 7072` | Autre port d'écoute (saisir alors l'URL dans la WebPart) |

Procédure de recette complète : [docs/RECETTE.md](docs/RECETTE.md).

## Passage au backend réel

1. Obtenir de l'équipe backend les informations de la section J du document d'intégration.
2. Renseigner `API_RESOURCE` et, si nécessaire, les routes dans `API_ROUTES`.
3. Ajouter `webApiPermissionRequests` dans `config/package-solution.json` et incrémenter la version.
4. Si le contrat diffère, adapter `parseEmbedConfig` / `parseReportList` dans `PowerBiApiService.ts` et les tests associés, puis mettre à jour `docs/INTEGRATION_API.md`.
5. `npm run build`, déployer `sharepoint/solution/power-bi-viewer.sppkg` dans le catalogue d'applications, approuver l'accès à l'API.
6. Ajouter la WebPart sur une page, renseigner l'URL du service et le rapport, puis dérouler la recette (section I du document d'intégration).

## Limites des tests sans accès à l'environnement du client

Ce qui est vérifié dans ce dépôt :

- compilation TypeScript, lint et build de production ;
- validation du contrat de réponse, normalisation des erreurs HTTP, calcul du délai de renouvellement ;
- requêtes émises vers l'API (routes, méthode, corps JSON) ;
- cycle de vie de la visionneuse (powerbi-client simulé) : intégration en lecture seule, renouvellement par `setAccessToken()`, événement `TokenExpired`, perte de droits, messages d'erreur, absence d'appels concurrents ;
- comportement HTTP du simulateur (CORS, 401, 400, 404, 403, 503, format d'erreur), vérifié avec des appels Power BI simulés ;
- absence du code de simulation dans le package de production.

À vérifier manuellement en local avec le simulateur ([docs/RECETTE.md](docs/RECETTE.md), partie A) : affichage d'un vrai rapport, interactions, renouvellement avec de vrais jetons, dernière version publiée, dimensionnement dans une page SharePoint.

Ce qui ne peut être validé que chez le client, avec l'API d'intégration réelle (partie B) :

- obtention du jeton par `AadHttpClient` (approbation, audience, scope) et CORS de l'API réelle ;
- accès d'un invité B2B et contrôle d'appartenance au groupe autorisé ;
- accès réseau depuis les postes des utilisateurs (proxy, domaines Power BI).
