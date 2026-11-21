# Construire le tableau de bord Power BI (≈ 30 min)

> Je n'ai pas pu générer de fichier `.pbix` : il faut Power BI Desktop (Windows) pour le créer et le valider.
> Les requêtes (`1_…m`) et mesures (`2_…dax`) sont prêtes à coller ; la mise en page ci-dessous est à poser une fois.

## 1. Données
1. Lancer l'agent une fois : `python -m mailagent run` → `output\powerbi_data\Demandes.csv` et `Parametres.csv`.
2. Power BI Desktop → *Transformer les données* → créer le paramètre `DossierDonnees`, puis les requêtes `Demandes` et
   `Parametres` (`1_Requetes_PowerQuery.m`). Fermer et appliquer.
3. Modèle : créer la table `Calendrier` et les mesures (`2_Mesures.dax`) ; relation `Calendrier[Date] → Demandes[Date_reception]`.
   Mettre `Demandes[ID]` et `Parametres` en « masquer dans la vue rapport ».

## 2. Pages
| Page | Visuels |
|---|---|
| **Vue d'ensemble** | Cartes : *Demandes reçues*, *Demandes ouvertes*, *Demandes en alerte*, *Sans réponse au-delà du SLA*, *Délai moyen de réponse (j)*, *Taux de réponse dans le SLA*. Histogramme empilé *Categorie* × *Statut*. Courbe/colonnes par `Calendrier[Annee_semaine]` (reçues). Anneau *Priorite*. Segments : Categorie, Statut, Date. |
| **À traiter** | Tableau filtré `Ouverte = 1` : Recu_le, Expediteur, Objet, Resume, Action_attendue, Priorite, Echeance, Age_jours, Alerte. Tri `Age_jours` décroissant. Colonne *Alerte* : Mise en forme conditionnelle → Couleur d'arrière-plan → *Valeur du champ* → mesure `Couleur alerte`. |
| **Tri par objet** | Matrice : lignes `Categorie` > `Objet_nettoye` ; valeurs *Demandes reçues*, *Demandes ouvertes*. (Équivalent du « tri par objet ».) |
| **Hebdo** | Segment `Annee_semaine` + les cartes ; sert de base au rapport envoyé le lundi. |

## 3. Alerting et diffusion (service Power BI)
* **Alertes de données** : épingler la carte *Demandes en alerte* (ou *Sans réponse au-delà du SLA*) à un tableau de bord →
  « … » → *Gérer les alertes* → seuil « au-dessus de 0 » → e-mail/notification. Fonctionne uniquement sur cartes, jauges,
  KPI ; **pas dans Power BI Desktop**.
  [Doc](https://learn.microsoft.com/en-us/power-bi/create-reports/service-set-data-alerts)
* **Rapport hebdomadaire** : *S'abonner au rapport* → planification hebdomadaire, destinataires.
  [Doc](https://learn.microsoft.com/en-us/power-bi/collaborate-share/service-report-subscribe)
  (L'agent envoie en plus un e-mail HTML de synthèse le lundi, indépendant de Power BI.)
* **Les alertes « sans réponse à J+3 » par e-mail restent envoyées par l'agent** : Power BI ne sait pas envoyer un
  mail par demande, et n'est actualisé qu'à intervalle fixe.

## 4. Actualisation : le point qui piège
Les CSV sont locaux. Le service Power BI ne les lit que via une **passerelle de données** installée sur le PC
([doc](https://learn.microsoft.com/en-us/power-bi/connect-data/service-gateway-onprem)). Plus simple : faire écrire
`output\powerbi_data` dans un dossier **OneDrive/SharePoint synchronisé** (mettre `output_dir` dans ce dossier) et
utiliser la source *SharePoint* / *Dossier SharePoint* → actualisation planifiée sans passerelle
([doc](https://learn.microsoft.com/en-us/power-bi/connect-data/refresh-scheduled-refresh)).
Alignez la fréquence d'actualisation (max 8/jour en Pro) avec celle de l'agent (30 min) : le tableau de bord ne sera
pas « temps réel ».
Licences : partager/alerter/s'abonner nécessite Power BI Pro (ou PPU/Premium) pour l'auteur **et** les lecteurs.

## 5. Changer un statut
Power BI est en lecture seule. Après traitement manuel (ou si l'agent n'a pas vu la réponse) :
`python -m mailagent set-status <ID> "Clôturé" "note facultative"` (l'ID est dans le tableau *À traiter*).
Pour une équipe, la bonne solution est une liste SharePoint ou Dataverse éditée via Power Apps (voir README).
