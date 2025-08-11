using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using System.Linq;

namespace ComptaFlow.Services
{
    public class TCDGeneratorSage
    {
        public void GenererTCDAvecFeuilPrecedente(string cheminFichierExcel, string cheminSortie)
        {
            var workbookSource = new Workbook(cheminFichierExcel);
            var workbookFinal = new Workbook();
            workbookFinal.Worksheets.Clear();

            // ➤ Définir le style par défaut
            SetDefaultStyle(workbookFinal);

            int compteurFeuil = 1;

            foreach (Worksheet feuilleSource in workbookSource.Worksheets)
            {
                if (feuilleSource.Name.StartsWith("Feuil", StringComparison.OrdinalIgnoreCase))
                    continue;

                NettoyerEtRenommerFeuille(feuilleSource, ref compteurFeuil);

                Worksheet copieFeuille = workbookFinal.Worksheets.Add(feuilleSource.Name);
                copieFeuille.Copy(feuilleSource);

                string nomFeuilTCD = $"Feuil{compteurFeuil++}";
                Worksheet feuilleTCD = workbookFinal.Worksheets.Add(nomFeuilTCD);

                var plageDonnees = feuilleSource.Cells.MaxDisplayRange;
                if (plageDonnees == null || plageDonnees.RowCount == 0 || plageDonnees.ColumnCount == 0)
                    continue;

                string plageAdresse = $"'{feuilleSource.Name}'!{plageDonnees.Address}";

                try
                {
                    // ➤ Génération et configuration du TCD
                    GenererEtConfigurerTCD(feuilleTCD, plageAdresse);

                    // ➤ Déplacement de la feuille TCD avant la copie
                    int idxFeuilleCopie = workbookFinal.Worksheets.IndexOf(copieFeuille);
                    int idxFeuilleTCD = workbookFinal.Worksheets.IndexOf(feuilleTCD);
                    workbookFinal.Worksheets[idxFeuilleTCD].MoveTo(idxFeuilleCopie);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur TCD feuille '{feuilleSource.Name}' : {ex.Message}");
                }
            }

            // Sauvegarde le fichier généré
            workbookFinal.Save(cheminSortie);

            // Supprime la dernière feuille avec ClosedXML
            SupprimerDerniereFeuilleAvecClosedXml(cheminSortie);

            Console.WriteLine($"Fichier généré : {cheminSortie}");
        }

        private void SetDefaultStyle(Workbook workbook)
        {
            Style styleDefaut = workbook.CreateStyle();
            styleDefaut.Font.Name = "Calibri";
            styleDefaut.Font.Size = 11;
            workbook.DefaultStyle = styleDefaut;
        }

        private void GenererEtConfigurerTCD(Worksheet feuilleTCD, string plageAdresse)
        {
            int indexPivot = feuilleTCD.PivotTables.Add(plageAdresse, "A1", "PivotTable1");
            PivotTable pivotTable = feuilleTCD.PivotTables[indexPivot];

            pivotTable.ShowInCompactForm();

            // Ajout du filtre sur Journal
            pivotTable.AddFieldToArea(PivotFieldType.Page, "Journal");

            // Ajout des champs en lignes
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Date");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Libellé écriture");

            foreach (PivotField rowField in pivotTable.RowFields)
            {
                rowField.IsAutoSort = true;
                rowField.IsAscendSort = true;
            }

            // Ajout du champ en valeur
            int dataFieldIndex = pivotTable.AddFieldToArea(PivotFieldType.Data, "Montant signé (XAF)");
            pivotTable.DataFields[dataFieldIndex].Function = ConsolidationFunction.Sum;
            pivotTable.DataFields[dataFieldIndex].DisplayName = "Somme de Montant signé (XAF)";

            // Application des sous-totaux
            foreach (PivotField rowField in pivotTable.RowFields)
            {
                if (rowField.Name == "Date")
                {
                    rowField.SetSubtotals(PivotFieldSubtotalType.Sum, true);
                    rowField.ShowSubtotalAtTop = true;
                }
                else if (rowField.Name == "Libellé écriture")
                {
                    rowField.SetSubtotals(PivotFieldSubtotalType.None, true);
                }
            }

            pivotTable.ShowRowGrandTotals = true;
            pivotTable.ShowColumnGrandTotals = true;

            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Fixer le filtre Journal à "TRANSFERT"
            PivotField? filtreJournal = null;
            foreach (PivotField field in pivotTable.PageFields)
            {
                if (field.Name == "Journal")
                {
                    filtreJournal = field;
                    break;
                }
            }

            if (filtreJournal != null)
            {
                foreach (PivotItem item in filtreJournal.PivotItems)
                {
                    item.IsHidden = ((string)item.Value) != "TRANSFERT";
                }
            }
        }

        public void SupprimerDerniereFeuilleAvecClosedXml(string cheminFichier)
        {
            using var workbook = new XLWorkbook(cheminFichier);

            // Vérifie qu’il y a au moins une feuille
            if (workbook.Worksheets.Count > 0)
            {
                var derniereFeuille = workbook.Worksheets.Last();
                workbook.Worksheets.Delete(derniereFeuille.Name);
                workbook.Save(); // Écrase le fichier existant
                Console.WriteLine($"Dernière feuille '{derniereFeuille.Name}' supprimée.");
            }
            else
            {
                Console.WriteLine("Aucune feuille à supprimer.");
            }
        }

        private void NettoyerEtRenommerFeuille(Worksheet feuille, ref int compteur)
        {
            string nomOriginal = feuille.Name;
            string nomNettoye = Regex.Replace(nomOriginal, @"[\\\/\*\[\]\?:']", "_");

            if (nomNettoye.Length > 31)
                nomNettoye = nomNettoye.Substring(0, 31);

            if (nomNettoye != nomOriginal)
            {
                Console.WriteLine($"Renommage feuille '{nomOriginal}' en '{nomNettoye}' pour éviter erreurs");
                feuille.Name = nomNettoye;
            }
        }
    }
}
