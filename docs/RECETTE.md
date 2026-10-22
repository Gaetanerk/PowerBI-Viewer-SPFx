# Recette — Visionneuse Power BI

Procédure de vérification du cahier des charges de la WebPart.

- **Partie A — en local**, sur le tenant SharePoint de développement. La WebPart appelle le **simulateur de l'API d'intégration** (`tools/Start-BackendSimulator.ps1`), qui respecte le contrat attendu du backend et génère de vrais Embed Tokens : le rapport AdventureWorks Sales s'affiche réellement. Réalisable dès maintenant.
- **Partie B — dans l'environnement du client**, avec l'API d'intégration réelle et un utilisateur invité B2B. Réalisable lorsque l'API est disponible.

---

## Partie A — Recette locale avec le simulateur

### Prérequis

| Élément | Détail |
|---|---|
| Poste | Node.js 22 (>= 22.14), dépendances installées (`npm install`), PowerShell |
| SharePoint | Un tenant de développement donnant accès au workbench hébergé |
| Power BI | Un compte avec une licence Pro, PPU ou un essai gratuit (nécessaire pour créer un espace de travail) |

### Mise en place

1. **Terminal 1** : démarrer le simulateur (fenêtre de connexion Microsoft au premier lancement ; il crée l'espace de travail « Visionneuse Power BI - Tests » et y publie `dev/samples/AdventureWorks Sales.pbix`) :

   ```powershell
   ./tools/Start-BackendSimulator.ps1 -SimulateErrors
   ```

   Il affiche `Simulateur de l'API d'intégration à l'écoute sur http://localhost:7071` et la liste des rapports. Laisser ce terminal visible : chaque appel de la WebPart y est journalisé.
2. **Terminal 2** : `npm start`, puis ouvrir `https://<tenant>.sharepoint.com/_layouts/workbench.aspx`.
3. Ajouter la WebPart **Visionneuse Power BI**, ouvrir ses propriétés (icône crayon), activer **Mode simulation** (URL laissée vide) et choisir **AdventureWorks Sales**.

> Ne pas lancer `npm run build` pendant que `npm start` tourne : le build de production remplace les fichiers servis et fait disparaître le mode simulation jusqu'au prochain `npm start`.

### A1 — Paramétrage (cahier des charges § 6)

| # | Étapes | Résultat attendu | OK / KO |
|---|---|---|---|
| A1.1 | Ouvrir le panneau de propriétés | Uniquement : URL de l'API, Rapport Power BI, Hauteur automatique (+ interrupteur Mode simulation, propre au build de développement). Aucun panneau d'administration supplémentaire. | |
| A1.2 | À l'ouverture du panneau | Liste déroulante des rapports alimentée par `GET /api/reports` ; ligne `GET /api/reports 200` dans le terminal du simulateur. | |
| A1.3 | Mode simulation désactivé, saisir `http://exemple.fr` | Erreur « Saisissez une URL HTTPS valide, sans paramètres. » | |
| A1.4 | Saisir `https://exemple.azurewebsites.net/?code=abc` | Même erreur : une clé de fonction n'est jamais acceptée. | |
| A1.5 | Arrêter le simulateur (Ctrl+C), recharger la page, ouvrir le panneau | Champ de saisie libre avec « Liste des rapports indisponible… ». Relancer ensuite le simulateur. | |
| A1.6 | Ajouter une seconde WebPart sans la configurer | Message « La visionneuse n'est pas configurée. Modifiez la WebPart (icône crayon)… ». En **Aperçu** : « Aucun rapport n'est encore configuré… ». | |
| A1.7 | Menu du workbench **Données de composant WebPart** | Les propriétés ne contiennent que `apiBaseUrl`, `reportId`, `autoHeight`, `customHeight`, `mockMode` : aucun jeton ni secret. | |

### A2 — Échanges avec le backend (§ 7)

Ouvrir F12 › Réseau, filtre `localhost:7071`, puis recharger la page.

| # | Étapes | Résultat attendu | OK / KO |
|---|---|---|---|
| A2.1 | Observer les requêtes au chargement | `OPTIONS /api/embed` (204, préliminaire CORS), puis `POST /api/embed`. | |
| A2.2 | Détail de `POST /api/embed` | En-têtes `Authorization: Bearer …`, `Content-Type: application/json` ; corps `{"reportId":"<GUID>"}`. | |
| A2.3 | Réponse | 200, `Cache-Control: no-store`, JSON `reportId`, `embedUrl`, `embedToken`, `expiration` (date ISO avec `Z`). | |
| A2.4 | Terminal du simulateur | `POST /api/embed 200 jeton généré (AdventureWorks Sales, expire …)`, sans jeton affiché. | |

### A3 — Affichage du rapport (§ 3, § 8)

| # | Étapes | Résultat attendu | OK / KO |
|---|---|---|---|
| A3.1 | Choisir **AdventureWorks Sales** | Indicateur de chargement discret, puis rapport affiché dans la page sous le bandeau « Mode simulation », sans demande de connexion Power BI. | |
| A3.2 | Cliquer sur un élément d'un graphique | Filtrage croisé des autres visuels. | |
| A3.3 | Ouvrir le volet **Filtres**, modifier un filtre | Filtre appliqué. | |
| A3.4 | Changer de page du rapport | Navigation entre pages fonctionnelle. | |
| A3.5 | Survoler un visuel | Info-bulles affichées. | |
| A3.6 | Observer le pourtour du rapport | Aucun en-tête, barre d'outils ou bouton plein écran ajouté par la WebPart (le bandeau n'existe qu'en simulation). | |
| A3.7 | F12 › Éléments : inspecter le rapport | Une iframe `app.powerbi.com/reportEmbed…` créée dans le conteneur par `powerbi-client` (aucune iframe construite à la main). | |

### A4 — Surface d'affichage (§ 5)

| # | Étapes | Résultat attendu | OK / KO |
|---|---|---|---|
| A4.1 | Section à une colonne | Le rapport occupe toute la largeur de la section, sans marge ni bordure ajoutée. | |
| A4.2 | Déplacer la WebPart dans une section à 2 puis 3 colonnes, puis pleine largeur | Le rapport suit la largeur de la colonne ; la largeur des sections SharePoint n'est pas modifiée. | |
| A4.3 | Hauteur automatique, grande fenêtre | Proportions du rapport respectées, hauteur jamais supérieure à la zone visible de la page. | |
| A4.4 | Réduire la hauteur de la fenêtre | Le rapport se réduit pour rester entièrement visible : pas de double barre de défilement. | |
| A4.5 | F12 › mode responsive : largeurs 375, 768, 1366, 1920 px | Adaptation automatique, sans défilement horizontal. | |
| A4.6 | Désactiver **Hauteur automatique**, régler 600 px puis 1200 px | Hauteur appliquée immédiatement, rapport ajusté dans la zone. | |

### A5 — Renouvellement de l'Embed Token (§ 9)

Relancer le simulateur avec des jetons courts : `./tools/Start-BackendSimulator.ps1 -TokenLifetimeMinutes 10 -SimulateErrors`.

| # | Étapes | Résultat attendu | OK / KO |
|---|---|---|---|
| A5.1 | Recharger la page, appliquer un filtre dans le rapport | Un `POST /api/embed` au chargement. | |
| A5.2 | Attendre ~5 minutes | Un nouveau `POST /api/embed` (nouveau jeton généré, visible dans le terminal), **sans rechargement du rapport** : le filtre et la page courante sont conservés. | |
| A5.3 | Mettre l'onglet en arrière-plan plus de 5 minutes, puis revenir | Renouvellement au retour si nécessaire ; rapport toujours fonctionnel. | |
| A5.4 | Arrêter le simulateur, attendre l'expiration du jeton (10 min) | Nouvelles tentatives espacées, puis message « L'autorisation d'affichage du rapport a expiré et n'a pas pu être renouvelée. Rechargez la page. » | |
| A5.5 | Choisir **[Test] Droits retirés au renouvellement (403)**, attendre ~5 minutes | Rapport affiché, puis fermé au renouvellement avec « Vous n'êtes pas autorisé à consulter ce rapport. » et les informations support (code `USER_FORBIDDEN`, référence = `correlationId` du terminal). | |
| A5.6 | F12 › Application : Local Storage et Session Storage ; F12 › Console | Aucun jeton stocké ni affiché. | |

### A6 — Dernière version publiée (§ 3)

| # | Étapes | Résultat attendu | OK / KO |
|---|---|---|---|
| A6.1 | Dans `app.powerbi.com`, espace « Visionneuse Power BI - Tests », ouvrir le rapport › **Modifier** › changer le titre d'un visuel › **Enregistrer** | — | |
| A6.2 | Recharger la page du workbench, sans recompiler ni modifier les propriétés | La modification est visible. | |
| A6.3 | Variante Power BI Desktop : modifier `dev/samples/AdventureWorks Sales.pbix`, enregistrer, relancer le simulateur avec `-PbixPath "dev/samples/AdventureWorks Sales.pbix"`, recharger | Nouvelle version visible : le rapport est remplacé sous le même identifiant, la WebPart n'est pas reconfigurée. | |

### A7 — Gestion des erreurs du backend (§ 10)

Rapports `[Test]` du catalogue (simulateur lancé avec `-SimulateErrors`).

| # | Étapes | Résultat attendu | OK / KO |
|---|---|---|---|
| A7.1 | **[Test] Accès refusé (403)** | « Vous n'êtes pas autorisé à consulter ce rapport. » + code `REPORT_FORBIDDEN` et référence. | |
| A7.2 | **[Test] Rapport supprimé (404)** | « Le rapport demandé est introuvable ou n'est plus disponible. » | |
| A7.3 | **[Test] Service indisponible (503)** | « Le service de rapports est momentanément indisponible… » | |
| A7.4 | Arrêter le simulateur, recharger la page | « Le service de rapports est inaccessible… » | |
| A7.5 | Pour chaque cas | Aucun jeton dans le message ; la référence affichée correspond au `correlationId` du terminal. | |

Les autres cas (400, 401, réponse invalide, jeton renvoyé déjà expiré) sont couverts par les tests unitaires : `npx heft test`.

### A8 — Package de production et isolation du mode simulation (§ 12, § 15)

| # | Étapes | Résultat attendu | OK / KO |
|---|---|---|---|
| A8.1 | Arrêter `npm start`, lancer `npm run build` | Build et tests réussis ; package `sharepoint/solution/power-bi-viewer.sppkg` généré. | |
| A8.2 | Déployer le package dans le catalogue d'applications du tenant de développement, ajouter la WebPart sur une page moderne | WebPart disponible, message « non configurée ». | |
| A8.3 | Ouvrir ses propriétés | **Pas** d'interrupteur Mode simulation : la simulation est absente du package. | |
| A8.4 | Relancer `npm start` pour reprendre le développement | Mode simulation de nouveau disponible dans le workbench. | |

---

## Partie B — Recette dans l'environnement du client

Prérequis : API d'intégration déployée, valeurs Entra ID reçues et reportées dans `services/apiConfig.ts` et `config/package-solution.json`, package déployé, accès à l'API approuvé dans SharePoint (voir le README, « Passage au backend réel »).

La procédure détaillée, avec les moyens de diagnostic, se trouve en section I de [INTEGRATION_API.md](INTEGRATION_API.md#i-recette-et-validation). Points couverts :

| # | Vérification | Exigence |
|---|---|---|
| B1 | Connexion d'un invité B2B au site, sans connexion supplémentaire | § 2 |
| B2 | Appel authentifié à l'API avec un jeton Entra ID (`AadHttpClient`) | § 7 |
| B3 | Contrôle d'accès côté serveur (groupe autorisé, rapport autorisé) | § 2, § 7 |
| B4 | Réception de l'Embed Token, réponse conforme au contrat | § 7, § 8 |
| B5 | Affichage et interactions du rapport réel | § 3, § 8 |
| B6 | Renouvellement du jeton auprès de l'API, sans rechargement | § 9 |
| B7 | Dernière version publiée visible au rechargement | § 3 |
| B8 | Refus d'accès : utilisateur hors groupe (403), rapport hors liste, requête sans jeton (401), perte de droits en cours de session | § 9, § 10 |
| B9 | Catalogue des rapports dans le panneau, si l'API l'expose | § 6 |

---

## Couverture du cahier des charges

| Section | Vérifiée par |
|---|---|
| § 3 Comportement fonctionnel, mise à jour des rapports | A3, A6, B5, B7 |
| § 5 Affichage, surface maximale | A4 |
| § 6 Paramètres | A1, B9 |
| § 7 Communication avec l'API | A2, tests unitaires, B2 à B4 |
| § 8 Intégration powerbi-client | A3, tests unitaires |
| § 9 Renouvellement des jetons | A5, tests unitaires, B6, B8 |
| § 10 Interface et erreurs | A7, tests unitaires, B8 |
| § 12 Mode simulation isolé | A8 |
| Sécurité : aucun jeton stocké ni affiché | A1.7, A5.6, A7.5 |
