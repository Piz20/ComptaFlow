using Aspose.Cells;
using System;
using System.IO;

namespace ComptaFlow.Services
{
    public class TCDGeneratorSage
    {
        public void GenererTCDPourChaqueFeuille(string cheminFichierExcel, string cheminSortie)
        {
            var workbook = new Workbook(cheminFichierExcel);
            int feuilleInitialeCount = workbook.Worksheets.Count;

            // Création des TCD dans le même fichier
            for (int i = 0; i < feuilleInitialeCount; i++)
            {
                var feuilleSource = workbook.Worksheets[i];

                if (feuilleSource.Name.StartsWith("TCD")) continue;

                var plageDonnees = DetecterPlageDonnees(feuilleSource);
                if (string.IsNullOrWhiteSpace(plageDonnees)) continue;

                var nomFeuille = feuilleSource.Name;
                if (nomFeuille.Contains(" ") || nomFeuille.Contains("-"))
                    nomFeuille = $"'{nomFeuille}'";

                var nomFeuilleTCD = $"TCD {i + 1}";
                var feuilleTCD = workbook.Worksheets.Add(nomFeuilleTCD);

                string plage = $"{nomFeuille}!{plageDonnees}";

                try
                {
                    int ptIndex = feuilleTCD.PivotTables.Add(plage, "A3", $"Pivot_{i + 1}");
                    var pivot = feuilleTCD.PivotTables[ptIndex];

                    var range = feuilleSource.Cells.CreateRange(plageDonnees);
                    int rowTitle = range.FirstRow;
                    int colStart = range.FirstColumn;
                    int colCount = range.ColumnCount;

                    int colDate = -1;
                    int colLibelle = -1;
                    int colMontant = -1;

                    for (int col = colStart; col < colStart + colCount; col++)
                    {
                        string header = feuilleSource.Cells[rowTitle, col].StringValue.Trim();

                        if (header.Equals("Date", StringComparison.OrdinalIgnoreCase))
                            colDate = col;

                        if (header.Equals("Libellé écriture", StringComparison.OrdinalIgnoreCase))
                            colLibelle = col;

                        if (header.Equals("Montant signé (XAF)", StringComparison.OrdinalIgnoreCase))
                            colMontant = col;
                    }

                    if (colDate == -1 || colLibelle == -1 || colMontant == -1)
                    {
                        Console.WriteLine($"❌ Feuille '{feuilleSource.Name}' : colonnes obligatoires manquantes.");
                        continue;
                    }

                    pivot.RowFields.Add(pivot.BaseFields[feuilleSource.Cells[rowTitle, colDate].StringValue]);
                    pivot.RowFields.Add(pivot.BaseFields[feuilleSource.Cells[rowTitle, colLibelle].StringValue]);
                    pivot.DataFields.Add(pivot.BaseFields[feuilleSource.Cells[rowTitle, colMontant].StringValue]);
                    pivot.DataFields[0].Function = ConsolidationFunction.Sum;

                    pivot.RefreshData();
                    pivot.CalculateData();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Erreur création TCD feuille '{feuilleSource.Name}' : {ex.Message}");
                    continue;
                }
            }

            // Sauvegarde intermédiaire dans cheminSortie
            workbook.Save(cheminSortie);

            // --- Copier toutes les feuilles sauf la dernière dans un nouveau fichier ---

            var workbookFinal = new Workbook();
            while (workbookFinal.Worksheets.Count > 0)
                workbookFinal.Worksheets.RemoveAt(0);

            int totalFeuilles = workbook.Worksheets.Count;

            for (int i = 0; i < totalFeuilles - 1; i++)  // Toutes sauf la dernière
            {
                var feuille = workbook.Worksheets[i];
                int nouvelleFeuilleIndex = workbookFinal.Worksheets.Add();
                var nouvelleFeuille = workbookFinal.Worksheets[nouvelleFeuilleIndex];
                nouvelleFeuille.Copy(feuille);
                nouvelleFeuille.Name = NettoyerNomFeuille(feuille.Name);
            }

            string dossierSortie = Path.GetDirectoryName(cheminSortie) ?? "";
            string nouveauFichier = Path.Combine(dossierSortie, "TCD COMPLET SAGE.xlsx");

            workbookFinal.Save(nouveauFichier);

            Console.WriteLine($"📄 Nouveau fichier sans la dernière feuille sauvegardé : {nouveauFichier}");
        }

        private string NettoyerNomFeuille(string nom)
        {
            if (string.IsNullOrEmpty(nom))
                return "Feuille";

            if (nom.Length > 31)
                nom = nom.Substring(0, 31);

            char[] interdits = { '\\', '/', '?', '*', '[', ']', ':' };
            foreach (var c in interdits)
                nom = nom.Replace(c, '_');

            nom = nom.Trim('\'');

            if (string.IsNullOrWhiteSpace(nom))
                nom = "Feuille";

            return nom;
        }

        private string? DetecterPlageDonnees(Worksheet feuilleSource)
        {
            var plage = feuilleSource.Cells.MaxDisplayRange;
            if (plage == null || plage.RowCount < 2 || plage.ColumnCount < 2)
                return null;

            var debut = CellsHelper.CellIndexToName(plage.FirstRow, plage.FirstColumn);
            var fin = CellsHelper.CellIndexToName(plage.FirstRow + plage.RowCount - 1, plage.FirstColumn + plage.ColumnCount - 1);
            return $"{debut}:{fin}";
        }
    }
}
