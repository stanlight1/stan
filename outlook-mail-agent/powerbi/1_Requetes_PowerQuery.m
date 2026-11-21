// =============================================================================
//  Power BI Desktop > Accueil > Transformer les données > Nouvelle source > Requête vide
//  > Éditeur avancé. Créer 3 éléments (coller chaque bloc dans une requête distincte).
// =============================================================================

// ---- (A) PARAMÈTRE  « DossierDonnees »  (Gérer les paramètres > Nouveau, type Texte) ----
//   Valeur : le dossier powerbi_data de l'agent, ex. :
//   C:\agent\outlook-mail-agent\output\powerbi_data
//   Pour l'actualisation planifiée dans le service Power BI sans passerelle, utiliser plutôt
//   la source « SharePoint / OneDrive » (voir GUIDE) - ce paramètre n'est alors plus utilisé.


// ---- (B) REQUÊTE  « Demandes » ----
let
    Source   = Csv.Document(File.Contents(DossierDonnees & "\Demandes.csv"),
                 [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Entetes  = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Types    = Table.TransformColumnTypes(Entetes, {
                 {"Recu_le", type datetime}, {"Echeance", type date}, {"Repondu_le", type datetime},
                 {"Alerte_envoyee_le", type datetime}, {"Ouverte", Int64.Type},
                 {"Age_jours", Int64.Type}, {"Delai_reponse_jours", Int64.Type}}, "en-US"),
    DateRec  = Table.AddColumn(Types, "Date_reception", each Date.From([Recu_le]), type date),
    Semaine  = Table.AddColumn(DateRec, "Annee_semaine",
                 each Text.From(Date.Year(Date.StartOfWeek([Date_reception], Day.Monday)))
                      & "-S" & Text.PadStart(Text.From(Date.WeekOfYear([Date_reception], Day.Monday)), 2, "0"),
                 type text)
in
    Semaine


// ---- (C) REQUÊTE  « Parametres » ----
let
    Source  = Csv.Document(File.Contents(DossierDonnees & "\Parametres.csv"),
                [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Entetes = Table.PromoteHeaders(Source, [PromoteAllScalars = true])
in
    Entetes
