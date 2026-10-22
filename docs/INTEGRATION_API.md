# Contrat d'intégration — WebPart « Visionneuse Power BI » × API d'intégration Power BI

| | |
|---|---|
| **Objet** | Interfaces, paramètres et exigences que l'API d'intégration (Azure Function) doit respecter pour alimenter la WebPart SPFx de visionneuse Power BI Embedded. |
| **Destinataires** | Équipe backend développant l'API d'intégration, organisation cliente (exploitation SharePoint / Entra ID) |
| **Version du document** | 1.0, 21/10/2026, correspond à la version 1.0.0.0 du package SPFx |
| **Statut** | **Contrat provisoire.** Les routes, noms de champs et identifiants Entra ID ci-dessous sont la proposition de la WebPart. Ils doivent être confirmés ou amendés par l'équipe backend (voir [section J](#j-points-restant-à-confirmer-avec-léquipe-backend)). Les valeurs données en exemple sont fictives. |

Les comportements décrits ici sont ceux du code livré (`src/webparts/powerBiViewer/services/PowerBiApiService.ts` et `components/PowerBiViewer.tsx`). Toute évolution du contrat doit être répercutée des deux côtés.

---

## A. Architecture des échanges

```text
Utilisateur invité B2B
   │  1. Connexion à SharePoint Online (tenant de l'organisation cliente, compte invité Entra ID B2B)
   ▼
Page SharePoint ─ WebPart SPFx « Visionneuse Power BI »
   │  2. AadHttpClient obtient silencieusement un jeton Entra ID délégué (audience = API d'intégration)
   │  3. POST {URL}/api/embed  + Authorization: Bearer <jeton Entra ID>
   ▼
API d'intégration (Azure Function)
   │  4. Valide le jeton, vérifie le groupe autorisé et l'autorisation sur le rapport
   │  5. S'authentifie auprès de Power BI avec son principal de service
   │  6. Génère un Embed Token en lecture seule
   ▼
API REST Power BI / Fabric (capacité Fabric dédiée)
   │
   ▼
WebPart ← { reportId, embedUrl, embedToken, expiration }
   │  7. powerbi-client crée l'iframe et charge le rapport avec l'Embed Token
   ▼
Power BI Embedded (app.powerbi.com) affiche le rapport dans la page
```

| Composant | Responsabilités | Ne fait pas |
|---|---|---|
| **SharePoint Online / Entra ID (organisation cliente)** | Authentifie l'utilisateur invité. Délivre le jeton Entra ID destiné à l'API d'intégration, après approbation de la permission par un administrateur. | — |
| **WebPart SPFx** | Appelle l'API avec `AadHttpClient`, valide la réponse, transmet l'Embed Token à `powerbi-client`, renouvelle le jeton avant expiration, affiche les erreurs. | Aucune authentification Power BI, aucune génération de jeton, aucun secret, aucun contrôle d'autorisation (elle n'est pas une frontière de sécurité). |
| **API d'intégration** | Valide le jeton Entra ID, autorise l'utilisateur et le rapport, génère l'Embed Token avec le principal de service. | Ne renvoie jamais de secret ni de jeton Entra ID / Power BI du principal de service. |
| **Power BI Embedded** | Vérifie l'Embed Token et sert le rapport (dernière version publiée). | — |

En résumé : **SharePoint/Entra ID authentifie l'utilisateur, l'API d'intégration autorise et génère l'Embed Token, la WebPart ne fait qu'afficher.**

---

## B. Configuration Entra ID attendue

### B.1 Principe

La WebPart demande un jeton via `AadHttpClient` pour une **ressource** (l'API d'intégration). Ce jeton est émis par le **tenant de l'organisation cliente**, y compris pour un utilisateur invité B2B : l'API doit donc exister (au moins comme principal de service / application d'entreprise) dans le tenant client.

Le jeton est demandé par l'application Microsoft **« SharePoint Online Web Client Extensibility »**, à laquelle l'administrateur accorde la permission déléguée vers l'API d'intégration lors de l'approbation dans SharePoint.

### B.2 Paramètres à fournir par l'équipe backend

| Paramètre | Public / secret | Utilisation côté WebPart | Exemple (fictif) |
|---|---|---|---|
| Tenant ID de l'organisation cliente | Public | Aucune (sert à l'équipe backend pour valider `tid` / `iss`) | `11111111-2222-3333-4444-555555555555` |
| Application (client) ID de l'API | Public | Peut servir de ressource `AadHttpClient` | `aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee` |
| Application ID URI | Public | **Ressource `AadHttpClient`** (`API_RESOURCE`) | `api://aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee` |
| Nom de la ressource Entra ID (nom d'affichage de l'application d'entreprise) | Public | `webApiPermissionRequests[].resource` dans `package-solution.json` | `Power BI Embed API` |
| Scope délégué | Public | `webApiPermissionRequests[].scope` | `user_impersonation` |
| Mode de contrôle du groupe autorisé (rôle d'application ou groupe) | Public | Aucune | rôle `Reports.Read` attribué au groupe autorisé |
| URL de base de l'Azure Function | Public | Propriété « URL de l'API d'intégration (Azure Function) » de la WebPart | `https://func-powerbi-pbi.azurewebsites.net` |
| Secret / certificat du principal de service Power BI | **Secret** | **Jamais transmis** | — |
| Clés de fonction Azure (`?code=`) | **Secret** | **Jamais utilisées** (la WebPart refuse une URL avec paramètres) | — |

### B.3 Audience, émetteur et version du jeton

Le jeton reçu dépend de la configuration `requestedAccessTokenVersion` (manifeste de l'application API) :

| Revendication | Jeton v1 (par défaut) | Jeton v2 |
|---|---|---|
| `aud` | Application ID URI (ex. `api://aaaa…`) | Application (client) ID |
| `iss` | `https://sts.windows.net/{tenant client}/` | `https://login.microsoftonline.com/{tenant client}/v2.0` |
| `tid` | Tenant ID de l'organisation cliente (y compris pour les invités) | idem |
| `scp` | Scope(s) délégué(s), ex. `user_impersonation` | idem |
| `roles` | Rôles d'application attribués à l'utilisateur ou à ses groupes | idem |
| `oid` | Identifiant de l'objet **invité** dans le tenant client | idem |
| `azp` / `appid` | Application cliente SharePoint (attendu : `08e18876-6177-487e-b8b5-cf950c1e598c`, *à vérifier sur un jeton réel*) | idem |

Le backend doit valider au minimum : signature (clés JWKS du tenant), `exp`/`nbf`, `iss` et `tid` = tenant client, `aud` = valeur configurée, présence du scope délégué attendu dans `scp`, puis l'appartenance au groupe autorisé.

Pour identifier un invité, utiliser `oid` (stable) plutôt que `upn`, qui peut être absent ou au format `prenom.nom_domaine#EXT#@…`.

### B.4 Scope délégué et rôle `Reports.Read` : à distinguer

| Si `Reports.Read` est… | Déclaré dans | Apparaît dans | Déclaration SPFx |
|---|---|---|---|
| un **scope délégué** (`oauth2PermissionScopes`) | « Exposer une API » | `scp` | Oui, `webApiPermissionRequests` |
| un **rôle d'application** (`appRoles`, membres de type *User/Group*) | « Rôles d'application » | `roles` | Non : attribution aux utilisateurs ou au groupe autorisé dans l'application d'entreprise |

Configuration recommandée : **un scope délégué** (ex. `user_impersonation`) demandé par la WebPart **et** un **rôle d'application `Reports.Read` attribué au groupe autorisé**, contrôlé par l'équipe backend dans `roles`. L'attribution de groupes à des rôles d'application nécessite une licence Entra ID P1. Variante : revendication `groups` (attention à la limite de 200 groupes, au-delà de laquelle le jeton ne contient plus la liste) ou appel Microsoft Graph côté backend.

> **Important.** L'approbation dans SharePoint accorde le scope à l'application « SharePoint Online Web Client Extensibility » pour **tout le tenant** : toute solution SPFx installée chez le client peut alors obtenir un jeton pour l'API d'intégration. L'autorisation doit donc reposer sur **l'utilisateur** (rôle / groupe autorisé et liste des rapports), jamais sur le seul fait de recevoir un jeton valide.

Si l'équipe backend active « Affectation requise » sur l'application d'entreprise, un utilisateur non affecté ne peut pas obtenir de jeton (erreur `AADSTS50105`). La WebPart affiche alors le message d'échec d'authentification, et non le message « non autorisé ».

### B.5 Permissions à approuver dans SharePoint

Déclaration ajoutée dans `config/package-solution.json` dès réception des valeurs :

```json
"webApiPermissionRequests": [
  { "resource": "<Nom de l'application API d'intégration>", "scope": "<scope délégué>" }
]
```

Après déploiement du package, un administrateur SharePoint de l'organisation cliente approuve la demande dans *Centre d'administration SharePoint › Avancé › Accès à l'API*.

### B.6 CORS

L'API est appelée **depuis le navigateur**, à partir des pages SharePoint.

| Réglage | Valeur attendue |
|---|---|
| Origines autorisées | `https://<tenant-client>.sharepoint.com` (production). Ajouter temporairement l'origine du tenant de développement si l'équipe backend souhaite un test croisé. Pas de `*`. |
| Méthodes | `GET`, `POST`, `OPTIONS` |
| En-têtes de requête | `Authorization`, `Content-Type`, `Accept` |
| Identifiants (cookies) | Non utilisés : `Access-Control-Allow-Credentials` inutile |
| Requête préliminaire `OPTIONS` | Elle ne porte **pas** de jeton : elle doit répondre 2xx avec les en-têtes CORS, sans exiger d'authentification. |

Utiliser le CORS de la plateforme Azure Functions **ou** des en-têtes posés par le code, pas les deux. Si App Service Authentication (Easy Auth) est activé, vérifier qu'il ne bloque pas la requête préliminaire.

---

## C. Endpoints nécessaires

### Règles communes

- **URL de base** : saisie par l'organisation cliente dans la WebPart, HTTPS obligatoire, **sans `/api`**, sans paramètre de requête (pas de clé de fonction). Les routes ci-dessous sont ajoutées à cette base.
- **Authentification** : `Authorization: Bearer <jeton Entra ID délégué>`, ajouté automatiquement par `AadHttpClient`. Niveau d'autorisation de la fonction : *anonymous* (aucune clé), la sécurité reposant sur la validation du jeton.
- **En-têtes envoyés** : `Accept: application/json` ; pour `POST`, `Content-Type: application/json`.
- **Délai** : la WebPart abandonne une requête au bout de **30 secondes**.
- **Réponses** : JSON (`Content-Type: application/json`), encodage UTF-8. Les propriétés supplémentaires sont ignorées.
- **Cache** : aucune mise en cache côté WebPart. Les réponses contenant un Embed Token doivent porter `Cache-Control: no-store`.

### C.1 Catalogue des rapports — `GET /api/reports` (facultatif)

| | |
|---|---|
| Rôle | Alimenter la liste déroulante « Rapport Power BI » du panneau de propriétés (utilisé par les administrateurs de l'organisation cliente au moment de la configuration de la page). |
| Méthode / route | `GET {URL}/api/reports` |
| Authentification | Jeton Entra ID de l'utilisateur qui configure la page |
| Paramètres | Aucun |
| Réponse 200 | **Tableau JSON** (non enveloppé) d'objets `{ id, name }` |
| Filtrage attendu | Ne renvoyer que les rapports que l'utilisateur appelant est autorisé à consulter. |

Exemple :

```http
GET https://func-powerbi-pbi.azurewebsites.net/api/reports
Authorization: Bearer eyJ0eXAiOi...
Accept: application/json
```

```json
[
  { "id": "rapport-001", "name": "Suivi des prestations" },
  { "id": "rapport-002", "name": "Rapport financier" }
]
```

| Champ | Type | Obligatoire | Description |
|---|---|---|---|
| `id` | string non vide | Oui | Identifiant que la WebPart renverra tel quel dans `POST /api/embed`. GUID Power BI ou alias propre à l'API. |
| `name` | string non vide | Oui | Libellé affiché dans la liste. |

Comportement de la WebPart : si l'endpoint n'existe pas (404), échoue ou renvoie un format invalide, le panneau propose une **saisie manuelle** de l'identifiant. Un tableau vide est accepté (saisie manuelle proposée).

Erreurs possibles : 401, 403, 404 (catalogue non implémenté), 500, 503 (format en [section F](#f-gestion-des-erreurs)).

### C.2 Informations d'intégration d'un rapport — `POST /api/embed`

| | |
|---|---|
| Rôle | Fournir de quoi afficher un rapport : appelé à **chaque chargement de page** par chaque visiteur. |
| Méthode / route | `POST {URL}/api/embed` |
| Authentification | Jeton Entra ID de l'utilisateur consultant la page |
| Corps de requête | `{ "reportId": "<identifiant configuré dans la WebPart>" }` |
| Réponse 200 | Objet `{ reportId, embedUrl, embedToken, expiration }` (voir [section D](#d-format-des-données)) |

Exemple :

```http
POST https://func-powerbi-pbi.azurewebsites.net/api/embed
Authorization: Bearer eyJ0eXAiOi...
Accept: application/json
Content-Type: application/json

{ "reportId": "rapport-001" }
```

```json
{
  "reportId": "f6bfd646-b718-44dc-a378-b73e6b528204",
  "embedUrl": "https://app.powerbi.com/reportEmbed?reportId=f6bfd646-b718-44dc-a378-b73e6b528204&groupId=be8908da-da25-452e-b220-163f52476cdd&w=2&config=eyJjbHVzdGVyVXJsIjoi...",
  "embedToken": "H4sIAAAAAAAEACWWxQ7rCBZE...",
  "expiration": "2026-12-01T15:00:00Z"
}
```

Traitements attendus côté backend :

1. Valider le jeton Entra ID ([B.3](#b3-audience-émetteur-et-version-du-jeton)) et l'appartenance au groupe autorisé.
2. Valider `reportId` (présent, chaîne, longueur raisonnable) ; sinon **400**.
3. Résoudre `reportId` dans la **liste des rapports autorisés côté serveur** (espace de travail + rapport Power BI) ; inconnu → **404**, non autorisé pour cet utilisateur → **403**.
4. Récupérer l'`embedUrl` du rapport (API REST Power BI `GET /groups/{groupId}/reports/{reportId}`).
5. Générer l'Embed Token (`POST https://api.powerbi.com/v1.0/myorg/GenerateToken`), limité à ce rapport et à son jeu de données, en **lecture seule** (`accessLevel: "View"`, pas d'édition ni d'enregistrement).
6. Répondre 200 avec `Cache-Control: no-store`.

Erreurs possibles : 400, 401, 403, 404, 500, 503.

### C.3 Renouvellement de l'Embed Token — même endpoint

La WebPart renouvelle le jeton en **rappelant `POST /api/embed` avec le même corps**. Aucun endpoint dédié n'est requis. Le backend doit refaire l'intégralité des contrôles (jeton Entra ID, groupe autorisé, rapport autorisé) à chaque appel.

Contrainte : le `reportId` de la réponse doit être **identique** à celui de la réponse initiale ; sinon la WebPart ferme le rapport (réponse invalide). L'`embedUrl` peut changer ; seule la valeur initiale est utilisée.

Si l'équipe backend préfère un endpoint distinct (ex. `POST /api/embed/token`), il faudra adapter la WebPart : le signaler en [section J](#j-points-restant-à-confirmer-avec-léquipe-backend).

### C.4 Simulateur de référence

Le dépôt de la WebPart contient un simulateur local de ce contrat, [tools/Start-BackendSimulator.ps1](../tools/Start-BackendSimulator.ps1), utilisé pour la recette sans l'API réelle. Il illustre le comportement attendu : routes, formats JSON, en-tête `Authorization` exigé, contrôle du `reportId` côté serveur, CORS et requête préliminaire `OPTIONS`, `Cache-Control: no-store`, format d'erreur, génération d'un Embed Token en lecture seule à chaque appel.

Il ne constitue **pas** une implémentation de production : il ne valide pas la signature du jeton Entra ID, ne contrôle aucun groupe et utilise un compte utilisateur Power BI au lieu d'un principal de service.

---

## D. Format des données

### D.1 Requête `POST /api/embed`

```ts
interface EmbedRequest {
  reportId: string; // identifiant configuré dans la WebPart (id du catalogue ou saisie manuelle)
}
```

### D.2 Réponse `POST /api/embed`

```ts
interface EmbedResponse {
  reportId: string;   // GUID du rapport Power BI
  embedUrl: string;   // URL d'intégration renvoyée par l'API REST Power BI
  embedToken: string; // Embed Token Power BI
  expiration: string; // date ISO 8601 avec fuseau
}
```

| Champ | Obligatoire | Règle de validation appliquée par la WebPart |
|---|---|---|
| `reportId` | Oui | **GUID du rapport Power BI** (`xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`). Si l'API utilise des alias dans la requête, la réponse doit contenir le vrai GUID : c'est lui que `powerbi-client` utilise. |
| `embedUrl` | Oui | URL `https://` dont l'hôte est `powerbi.com` ou un sous-domaine (ex. `app.powerbi.com`). Toute autre URL est rejetée. |
| `embedToken` | Oui | Chaîne non vide. Jamais affichée, journalisée ni stockée par la WebPart (ni `localStorage`, ni `sessionStorage`, ni propriétés de page). |
| `expiration` | Oui | **ISO 8601 avec fuseau obligatoire** : `2026-12-01T15:00:00Z` (recommandé, UTC) ou `2026-12-01T16:00:00+01:00`. Secondes et millisecondes facultatives. Une date sans fuseau (`2026-12-01T15:00:00`) est **rejetée**. Le champ `expiration` renvoyé par l'API GenerateToken peut être retransmis tel quel. |

Toute réponse 200 non conforme est traitée comme une **réponse invalide** : le rapport n'est pas affiché.

### D.3 Réponse `GET /api/reports`

```ts
type ReportsResponse = Array<{ id: string; name: string }>;
```

### D.4 Corps d'erreur

```ts
interface ErrorResponse {
  error: {
    code: string;           // code stable, en majuscules (ex. REPORT_FORBIDDEN)
    message: string;        // message technique, non affiché à l'utilisateur
    correlationId?: string; // identifiant permettant de retrouver la requête dans les journaux de l'API
  };
}
```

---

## E. Sécurité et accès B2B

- La WebPart utilise exclusivement **`AadHttpClient`**. Le jeton Entra ID est obtenu silencieusement par SharePoint pour l'utilisateur connecté, sans connexion interactive supplémentaire.
- **Le backend valide chaque jeton Entra ID reçu** : signature, expiration, émetteur et tenant client, audience, scope délégué.
- **Le backend vérifie l'autorisation de l'utilisateur B2B** : rôle ou appartenance au groupe autorisé, puis droit sur le rapport demandé.
- **Le backend génère les Embed Tokens avec son principal de service** ; ses identifiants (secret, certificat) restent dans l'Azure Function (idéalement Key Vault ou identité managée). **Aucun secret, aucun identifiant de service, aucun jeton Entra ID ou Power BI du principal de service n'est transmis à la WebPart.**
- **Ne pas se fier au `reportId` envoyé par le navigateur** : il est modifiable par l'utilisateur. Le serveur le résout dans sa propre liste de rapports autorisés et refuse tout identifiant hors liste, y compris un GUID Power BI valide.
- **Portée minimale de l'Embed Token** : un seul rapport et son jeu de données, `accessLevel: "View"`, durée de vie raisonnable. Si une sécurité au niveau des lignes (RLS) est requise, l'identité effective doit être dérivée du jeton Entra ID côté serveur, jamais d'un paramètre client.
- La WebPart intègre le rapport avec `permissions: Read` ; seule la portée de l'Embed Token constitue toutefois une garantie de sécurité.
- **Accessibilité réseau** : l'API est appelée depuis le navigateur des utilisateurs de l'organisation invitée. Elle doit être joignable **depuis Internet, en HTTPS, sans VPN**, si les utilisateurs finaux n'en disposent pas. Les postes doivent également accéder aux domaines Power BI (`app.powerbi.com` et domaines associés, selon la liste officielle Microsoft des URL Power BI).
- **Journalisation** : ne jamais journaliser les jetons (Entra ID ou Embed Token). Journaliser `correlationId`, `oid`, `tid`, `reportId`, décision d'autorisation, statut HTTP et durée.
- Les champs `error.code` et `error.correlationId` sont **affichés à l'utilisateur** (rubrique « Informations pour le support ») : ils ne doivent contenir aucune donnée sensible.

---

## F. Gestion des erreurs

### F.1 Format

Toute réponse d'erreur porte le corps JSON défini en [D.4](#d4-corps-derreur) :

```json
{
  "error": {
    "code": "REPORT_FORBIDDEN",
    "message": "User is not allowed to view this report",
    "correlationId": "4f1c2b7e-9a3d-4c8e-8f61-0d2e5b7a9c10"
  }
}
```

Si le corps est absent ou illisible, la WebPart s'appuie sur le seul statut HTTP.

### F.2 Statuts et comportement de la WebPart

| Statut | Cas d'usage côté backend | Code suggéré | Message affiché par la WebPart | Pendant un renouvellement |
|---|---|---|---|---|
| **400** | Corps absent, `reportId` manquant ou mal formé | `INVALID_REQUEST` | « La demande d'affichage du rapport est invalide… » | Rapport fermé |
| **401** | Jeton absent, expiré, signature, audience ou tenant invalide | `UNAUTHENTICATED` | « Votre identité n'a pas pu être vérifiée… » | Rapport fermé |
| **403** | Utilisateur hors groupe autorisé, ou rapport non autorisé pour cet utilisateur | `USER_FORBIDDEN` / `REPORT_FORBIDDEN` | « Vous n'êtes pas autorisé à consulter ce rapport. » | Rapport fermé |
| **404** | `reportId` inconnu de la liste serveur, ou rapport supprimé dans Power BI | `REPORT_NOT_FOUND` | « Le rapport demandé est introuvable… » | Rapport fermé |
| **500** | Erreur interne, échec d'appel à Power BI | `INTERNAL_ERROR` / `POWERBI_ERROR` | « Le service de rapports est momentanément indisponible… » | Nouvelles tentatives |
| **503** | Indisponibilité temporaire (capacité, limitation) | `SERVICE_UNAVAILABLE` | idem 500 | Nouvelles tentatives |

Autres cas :

- 408 et 429 sont traités comme 500/503 (nouvelles tentatives lors d'un renouvellement) ; les autres statuts 4xx comme 400.
- Échec réseau, CORS refusé ou délai de 30 s dépassé : « Le service de rapports est inaccessible… ».
- Échec d'obtention du jeton Entra ID (permission non approuvée, ressource inconnue, `AADSTS…`) : message d'authentification ; le code `AADSTS` est affiché en information support.
- Réponse 200 non conforme : « Le service de rapports a renvoyé une réponse inattendue… ».

Pour le chargement initial, aucune nouvelle tentative automatique : l'utilisateur recharge la page.

---

## G. Renouvellement des tokens

| Point | Comportement de la WebPart | Attente vis-à-vis de l'API |
|---|---|---|
| Date d'expiration | Lue dans `expiration` (date ISO 8601 avec fuseau). Comparée à l'horloge du poste. | Renvoyer l'expiration réelle de l'Embed Token (champ `expiration` de GenerateToken). Durée de vie recommandée : **≥ 15 minutes** (60 minutes usuelles). |
| Déclenchement | **5 minutes avant l'expiration**. Si la durée de vie restante est inférieure à 5 min 10 s : à mi-durée, jamais moins de 10 s après la réception. | — |
| Mécanisme | Nouvel appel `POST /api/embed` avec le même `reportId`, puis `report.setAccessToken()` : **pas de rechargement** du rapport (filtres et page courante conservés). | Générer un nouvel Embed Token à chaque appel, avec les mêmes contrôles qu'au chargement. |
| Appels successifs | Un renouvellement par période de validité, par WebPart affichée. Aucune requête concurrente pour une même WebPart. Plusieurs WebParts sur une page = appels indépendants. Une page restée ouverte une journée enchaîne les renouvellements. | Accepter un nombre illimité d'appels successifs pour un même utilisateur et un même rapport (prévoir la limitation éventuelle en conséquence). |
| Onglet en arrière-plan / poste en veille | Au retour sur l'onglet, renouvellement immédiat si l'expiration est à moins de 5 minutes. Si le jeton a déjà expiré, le rapport est rechargé après `setAccessToken()`. L'événement Power BI `TokenExpired` déclenche aussi un renouvellement. | — |
| Erreur transitoire (réseau, 408, 429, 5xx) | Nouvelles tentatives après 15 s, 30 s, 60 s… tant que le jeton courant est valide. À l'expiration : « L'autorisation d'affichage du rapport a expiré et n'a pas pu être renouvelée. Rechargez la page. » | Renvoyer 503 plutôt que 500 pour une indisponibilité temporaire. |
| Perte de droits | Réponse 401 / 403 / 404 / 400 ou réponse invalide : le rapport est **immédiatement fermé** et le message correspondant affiché. | Retirer l'accès dès que l'utilisateur n'est plus autorisé : **403**. L'Embed Token déjà délivré reste techniquement valide jusqu'à son expiration : la durée de vie du jeton borne le délai de prise en compte d'un retrait de droits. |

---

## H. Configuration et déploiement chez le client

| # | Action | Responsable |
|---|---|---|
| 1 | Enregistrer l'application API dans le tenant client (ou provisionner son principal de service), exposer le scope délégué, définir le rôle `Reports.Read`, l'attribuer au groupe autorisé. | Équipe backend + administrateur Entra ID de l'organisation cliente |
| 2 | Déployer l'Azure Function : niveau *anonymous* + validation du jeton, CORS ([B.6](#b6-cors)), accès Internet HTTPS, principal de service autorisé sur les espaces de travail Power BI (capacité Fabric dédiée). | Équipe backend |
| 3 | Transmettre les valeurs de la [section B.2](#b2-paramètres-à-fournir-par-léquipe-backend) et confirmer les points de la [section J](#j-points-restant-à-confirmer-avec-léquipe-backend). | Équipe backend |
| 4 | Renseigner la ressource (`API_RESOURCE`) et `webApiPermissionRequests`, générer le package `power-bi-viewer.sppkg`. | Équipe SPFx |
| 5 | Déployer le package dans le catalogue d'applications SharePoint du tenant. | Administrateur SharePoint de l'organisation cliente |
| 6 | Approuver la demande d'accès à l'API (*Centre d'administration SharePoint › Avancé › Accès à l'API*). | Administrateur SharePoint / Entra ID de l'organisation cliente |
| 7 | Sur la page de diffusion : ajouter la WebPart « Visionneuse Power BI », renseigner l'URL de l'Azure Function, sélectionner le rapport, choisir la hauteur, publier la page. | Organisation cliente |
| 8 | Vérifier l'affichage avec un compte invité B2B membre du groupe autorisé, puis avec un compte non membre (recette, section I). | Organisation cliente + équipe backend |

---

## I. Recette et validation

Prérequis : un compte invité B2B **membre** du groupe autorisé, un compte invité **non membre**, un rapport autorisé, un rapport hors liste, accès aux journaux de l'Azure Function.

| # | Étape | Résultat attendu | Diagnostic |
|---|---|---|---|
| 1 | Connexion de l'invité B2B au site SharePoint | Page affichée sans connexion supplémentaire. | — |
| 2 | Appel authentifié à l'Azure Function | Outils de développement du navigateur (F12 › Réseau) : `POST …/api/embed` avec en-tête `Authorization: Bearer …`, préliminaire `OPTIONS` en 2xx. | Échec avant envoi = CORS, réseau ou jeton Entra ID (message d'authentification, code `AADSTS` éventuel). |
| 3 | Contrôle d'accès par le backend | Journal de l'API : `tid` l'organisation cliente, `oid` de l'invité, rôle / groupe autorisé reconnu, rapport autorisé. | Corréler avec le `correlationId`. |
| 4 | Récupération de l'Embed Token | Réponse 200 conforme à [D.2](#d2-réponse-post-apiembed), `expiration` avec fuseau. | Réponse invalide = message « réponse inattendue ». Vérifier `reportId` (GUID) et `embedUrl` (powerbi.com). |
| 5 | Affichage du rapport | Rapport visible dans SharePoint, occupant la largeur de la section, sans double barre de défilement. | Console : lignes `[PowerBiViewer] Erreur Power BI` avec `errorCode` et `requestId` (sans jeton). |
| 6 | Interactions | Filtres, segments, info-bulles, exploration et navigation entre pages fonctionnels. | — |
| 7 | Renouvellement du jeton | Pour le test, l'équipe backend réduit temporairement la durée de vie (ex. 10 min) : nouvel appel `POST /api/embed` ~5 min avant l'expiration, sans rechargement visible (filtres et page conservés). | Onglet Réseau : un seul appel par renouvellement. |
| 8 | Dernière version publiée | Modifier puis republier le rapport dans Fabric (même rapport, remplacement) ; recharger la page SharePoint : modification visible sans action SPFx. | Si non visible : vérifier que l'identifiant du rapport n'a pas changé. Les données dépendent par ailleurs de l'actualisation du modèle sémantique. |
| 9a | Refus : invité non membre de groupe autorisé | Réponse 403, message « Vous n'êtes pas autorisé à consulter ce rapport. » | Journal de l'API : refus motivé. |
| 9b | Refus : rapport hors liste (`reportId` modifié) | Réponse 403 ou 404, rapport non affiché. | — |
| 9c | Refus : requête sans jeton (outil HTTP) | Réponse 401. | — |
| 9d | Perte de droits en cours de session | Retirer l'utilisateur du groupe autorisé ; au renouvellement suivant, réponse 403 et fermeture du rapport. | Délai dépendant de la durée de vie des jetons et de la propagation Entra ID. |

**Diagnostic sans divulgation de jetons** : ne jamais copier un en-tête `Authorization` ni un `embedToken` dans un ticket ou un courriel. Échanger uniquement l'heure, l'URL de la page, le statut HTTP, `error.code`, `correlationId` (affichés dans la WebPart) et, côté Power BI, le `requestId` de la console. Pour inspecter les revendications d'un jeton Entra ID, le faire localement, sans le transmettre.

---

## J. Points restant à confirmer avec l'équipe backend

| # | Sujet | Valeur provisoire utilisée par la WebPart | Impact si différent |
|---|---|---|---|
| 1 | Tenant ID de l'organisation cliente | — | Validation `tid` / `iss` côté backend |
| 2 | Application (client) ID de l'API | — | Documentation, éventuellement ressource `AadHttpClient` |
| 3 | Application ID URI (ressource `AadHttpClient`) | `api://00000000-0000-0000-0000-000000000000` (**fictif**) | Modification de `apiConfig.ts` + nouveau package |
| 4 | Nom de la ressource Entra ID (application d'entreprise) | — | `webApiPermissionRequests` + nouveau package |
| 5 | Scope délégué | `user_impersonation` (**hypothèse**) | `webApiPermissionRequests` + nouveau package |
| 6 | Nature de `Reports.Read` (scope ou rôle d'application) et mode de contrôle groupe autorisé | Rôle d'application attribué au groupe autorisé (**recommandation**) | Configuration Entra ID ; aucun impact code si le scope délégué est inchangé |
| 7 | Enregistrement de l'API : tenant client ou multi-tenant du fournisseur de l'API | Principal de service présent dans le tenant client (**obligatoire**) | Faisabilité de l'approbation SharePoint |
| 8 | Version des jetons (v1 / v2) et valeur de `aud` attendue | — | Validation côté backend uniquement |
| 9 | URL de base de l'Azure Function | Saisie dans la WebPart | Aucun impact code |
| 10 | Routes | `GET /api/reports`, `POST /api/embed` | Modification de `API_ROUTES` + nouveau package |
| 11 | Disponibilité du catalogue `GET /api/reports` | Facultatif (saisie manuelle sinon) | Aucun impact code |
| 12 | Format de l'identifiant de rapport en requête (GUID Power BI ou alias) | Chaîne libre ; `reportId` de réponse = GUID Power BI | Validation de réponse à adapter si le GUID ne peut pas être renvoyé |
| 13 | Noms et formats des champs de réponse | `reportId`, `embedUrl`, `embedToken`, `expiration` (ISO 8601 avec fuseau) | Modification de `parseEmbedConfig` + nouveau package |
| 14 | Endpoint de renouvellement | Réutilisation de `POST /api/embed` | Modification du service si endpoint distinct |
| 15 | Format d'erreur et liste des codes | `{ error: { code, message, correlationId } }`, codes de la [section F](#f-gestion-des-erreurs) | Affichage des informations support uniquement |
| 16 | Durée de vie des Embed Tokens | ≥ 15 min recommandée | Fréquence des renouvellements, délai de retrait des droits |
| 17 | Sécurité au niveau des lignes (RLS) | Non prise en compte (aucun impact WebPart) | Côté backend uniquement |
| 18 | Origines CORS autorisées | `https://<tenant-client>.sharepoint.com` | Configuration Azure Function |
| 19 | Exposition réseau de l'API (Internet public, Front Door, APIM, point de terminaison privé) | Joignable depuis Internet sans VPN | Bloquant si l'API n'est pas joignable par les postes de l'organisation invitée |
| 20 | Domaines Power BI autorisés par le proxy de l'organisation invitée | — | Bloquant pour l'affichage du rapport |
