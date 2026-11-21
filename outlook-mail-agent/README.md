# Agent de suivi des demandes Outlook → tableau de bord (sans licence)

Agent Python (Windows) qui relève les mails reçus sur **une adresse donnée**, les trie par objet, les résume (IA),
alimente une **base de suivi + un tableau de bord HTML** (Power BI Desktop en option), prévient quand **aucune réponse n'a été faite après 3 jours**
et envoie un **reporting hebdomadaire**.

| Besoin | Où / comment |
|---|---|
| Récupérer les mails d'une adresse (à renseigner) | `target_address` dans `config.yaml` ; sources `outlook_com` ou `graph` |
| Trier par objet | colonnes `Categorie` (règles `categories:`) et `Objet_nettoye` ; page « Tri par objet » de Power BI |
| Résumé + tracking + tableau de bord + alertes | base `output/suivi.db` → **`output/tableau_de_bord.html`** (autonome, partageable) ; CSV `output/powerbi_data/` pour Power BI Desktop (option) |
| Alerte « sans retour après 3 jours » | e-mail de l'agent aux `alerts.recipients`, **une seule fois par demande** ; `sla.days` |
| Reporting hebdo | e-mail HTML de l'agent (lundi 08:00) + abonnement Power BI (optionnel) |

**Architecture** : `Outlook → agent (Python) → SQLite → tableau_de_bord.html (+ CSV pour Power BI Desktop)`.

**Pourquoi HTML et pas Power BI ?** Sans licence Pro, Power BI ne permet ni partage, ni abonnement, ni alertes, ni actualisation
planifiée : le rapport ne serait lisible que par son auteur, sur son PC. Le HTML est généré par l'agent à chaque passage, s'ouvre
dans n'importe quel navigateur (filtre, tri de colonnes, groupement par objet), sans licence ni serveur. Placez `output_dir` sur un
dossier réseau/OneDrive/SharePoint pour le partager. Limite : c'est un instantané statique (F5 pour actualiser), pas un outil
d'exploration libre comme Power BI. Les alertes et le reporting par e-mail sont envoyés par l'agent, donc sans licence non plus.

## Installation (Windows)
```powershell
cd outlook-mail-agent
py -m venv .venv ; .\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
copy config.example.yaml config.yaml        # puis éditer : target_address, destinataires, catégories
setx ANTHROPIC_API_KEY "sk-ant-..."         # optionnel : résumés par IA (sinon heuristique)
python -m mailagent run -c config.yaml      # 1er passage manuel
powershell -ExecutionPolicy Bypass -File .\install_tasks.ps1   # planification (30 min + lundi 08:00)
```
Dashboard : ouvrir `output/tableau_de_bord.html`. Power BI Desktop (gratuit, usage personnel) : `powerbi/GUIDE_POWERBI.md`.
Essai sans Outlook : `python -m mailagent.demo` puis `source.type: json_folder` + `notifier.type: file`
(les mails d'alerte sont écrits dans `output/outbox` au lieu d'être envoyés).

## Fonctionnement
1. **Relève** : mails des `lookback_days` derniers jours dont À/Cc contient `target_address` ; doublons, newsletters et
   expéditeurs ignorés exclus. **Une demande = une conversation.**
2. **Analyse** : catégorie par règles sur l'objet (sans « RE:/TR: ») ; résumé, action attendue, priorité, échéance citée
   et criticité par Claude (`claude-haiku-5-5`), avec repli heuristique si pas de clé / pas de réseau.
3. **Réponse détectée** automatiquement si un mail est envoyé dans la même conversation (dossier Éléments envoyés)
   → statut « Répondu ». Sinon : `python -m mailagent set-status <ID> "Clôturé" "note"`.
4. **Alertes** (colonne *Alerte*) : `Critique`, `SLA dépassé`, `Échéance dépassée`, `Échéance proche`.
   Mail immédiat pour les demandes critiques ; mail « sans réponse » à J+3.
5. **CSV régénérés** à chaque passage (`Demandes.csv`, `Parametres.csv`) ; la base SQLite fait foi.

## Choix techniques à connaître (franchement)
* **`outlook_com` ne fonctionne qu'avec Outlook de bureau classique**, ouvert/lançable, session Windows ouverte.
  Le « nouvel Outlook » n'expose pas COM ([Microsoft](https://learn.microsoft.com/en-us/office/vba/outlook/concepts/getting-started/using-the-object-model)),
  et la tâche planifiée ne tourne pas PC verrouillé/éteint.
* **`graph` est la voie robuste** (serveur, sans Outlook, sans session) : enregistrement d'une app Azure AD avec
  permissions d'application `Mail.Read` + `Mail.Send`, **à restreindre à la seule boîte** par une
  [Application Access Policy](https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access) (sinon l'app lit toutes les boîtes).
  Nécessite un administrateur Entra/Exchange. Secret via `GRAPH_CLIENT_SECRET`.
* **Confidentialité** : avec l'IA activée, l'objet/corps des mails (tronqué à 4000 car.) part vers l'API Anthropic.
  Vérifiez la conformité RGPD/politique interne ; sinon `ai.enabled: false`.
* **Jours ouvrés** (`sla.business_days: true`, par défaut) : un mail du vendredi n'est « en retard » que le mercredi.
  Passez à `false` pour des jours calendaires stricts. Les jours fériés ne sont pas gérés (voir propositions).
* Une réponse = tout envoi dans la conversation (y compris un simple transfert) ; en cas de doute, corrigez le Statut.
* Testé ici : logique, base, CSV, alertes, rapports (`pytest`).
  **Non testé** : backends Outlook COM et Graph, et le rapport Power BI lui-même (aucun Power BI Desktop ici ; requêtes M et mesures DAX non exécutées).
  Faites un 1er essai avec `notifier.type: file`.

## Propositions pour aller plus loin (par ordre de valeur)
1. **Passer à Microsoft Graph** + hébergement sur un serveur/VM ou Azure Function : fiabilité 24/7.
2. **Si vous obtenez des licences Power BI Pro**, ou **Base partagée** (liste SharePoint, Dataverse ou SQL Azure) + **formulaire Power Apps** pour changer les statuts :
   Power BI s'y connecte en direct (actualisation plus fine), plusieurs personnes peuvent traiter les demandes.
3. **Détection des relances** : un nouveau mail du demandeur dans une conversation déjà « Répondue » doit la rouvrir
   (actuellement une conversation = une demande, la relance n'est pas détectée).
4. **Alerte progressive** (J+2 au responsable, J+3 au manager, J+5 escalade) et alertes Teams (webhook) plutôt que seulement mail.
5. **SLA par catégorie/priorité** (ex. Critique 4 h, Facturation 5 j) + **jours fériés** (lib `holidays`).
6. **Assignation** : colonne « Responsable » proposée par l'IA selon la catégorie, + brouillon de réponse généré
   (avec validation humaine obligatoire avant envoi).
7. **Accusé de réception automatique** au demandeur avec numéro de ticket ; numéro de ticket dans l'objet pour fiabiliser le suivi.
8. **Pièces jointes** : sauvegarde + extraction (PDF) avant résumé ; **anti-spam/phishing** avant traitement.
9. **Évaluation de la qualité IA** : relire un échantillon hebdomadaire de résumés/priorités, ajuster les règles et le prompt.
10. Si le volume dépasse ~quelques centaines de demandes/mois ou si plusieurs personnes traitent, un **vrai outil de ticketing**
    (Jira Service Management, Freshdesk, Zendesk, ou Power Automate + Planner) est plus adapté que ce montage maison.

## Tests
`python -m pytest -q`
