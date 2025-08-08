using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ComptaFlow.Services
{
    public class TCDGeneratorIntegra
    {
        public void GenererTCD(string cheminFichierExcel, string cheminSortie)
        {
            try
            {
                var workbookSource = new Workbook(cheminFichierExcel);
                var feuilleSource = workbookSource.Worksheets["Rapport Op_Trait"];

                if (feuilleSource == null)
                {
                    Console.WriteLine("❌ Feuille 'Rapport Op_Trait' introuvable.");
                    return;
                }

                var plageDonnees = feuilleSource.Cells.MaxDisplayRange;
                if (plageDonnees == null || plageDonnees.RowCount == 0 || plageDonnees.ColumnCount == 0)
                {
                    Console.WriteLine("❌ Plage de données vide.");
                    return;
                }

                var workbookFinal = new Workbook();
                workbookFinal.Worksheets.Clear();
                SetDefaultStyle(workbookFinal);

                Worksheet feuilleTCD = workbookFinal.Worksheets.Add("TCD_RapportOpTrait");

                string plageAdresse = $"'{feuilleSource.Name}'!{plageDonnees.Address}";
                int indexPivot = feuilleTCD.PivotTables.Add(plageAdresse, "A1", "Pivot_RapportOpTrait");
                PivotTable pivotTable = feuilleTCD.PivotTables[indexPivot];

                pivotTable.ShowInCompactForm();

                // ➤ Ajout des champs en filtre
                AjouterChampSiExiste(pivotTable, PivotFieldType.Page, "Franchisé");
                AjouterChampSiExiste(pivotTable, PivotFieldType.Page, "Agence");
                AjouterChampSiExiste(pivotTable, PivotFieldType.Page, "Service");

                // ➤ Ajout des champs en lignes
                AjouterChampSiExiste(pivotTable, PivotFieldType.Row, "Date");

                // ➤ Ajout du champ en valeur
                if (AjouterChampSiExiste(pivotTable, PivotFieldType.Data, "Montant", out int dataFieldIndex))
                {
                    pivotTable.DataFields[dataFieldIndex].Function = ConsolidationFunction.Sum;
                    pivotTable.DataFields[dataFieldIndex].DisplayName = "Somme de Montant";
                }
                else
                {
                    Console.WriteLine("❌ Le champ 'Montant' n’a pas pu être ajouté.");
                    return;
                }

                // ➤ Totaux
                pivotTable.ShowRowGrandTotals = true;
                pivotTable.ShowColumnGrandTotals = false;

                pivotTable.RefreshData();
                pivotTable.CalculateData();

                workbookFinal.Save(cheminSortie);
                Console.WriteLine($"✅ TCD généré avec succès : {cheminSortie}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Une erreur est survenue lors de la génération du fichier.");
                Console.WriteLine($"Détails : {ex.Message}");
            }
        }

        private void SetDefaultStyle(Workbook workbook)
        {
            Style styleDefaut = workbook.CreateStyle();
            styleDefaut.Font.Name = "Calibri";
            styleDefaut.Font.Size = 11;
            workbook.DefaultStyle = styleDefaut;
        }

        private bool AjouterChampSiExiste(PivotTable pivotTable, PivotFieldType type, string nomChamp)
        {
            foreach (PivotField field in pivotTable.BaseFields)
            {
                if (field.Name.Trim().Equals(nomChamp.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    pivotTable.AddFieldToArea(type, nomChamp);
                    return true;
                }
            }
            Console.WriteLine($"⚠️ Champ '{nomChamp}' non reconnu par le TCD.");
            return false;
        }

        private bool AjouterChampSiExiste(PivotTable pivotTable, PivotFieldType type, string nomChamp, out int indexAjoute)
        {
            indexAjoute = -1;
            foreach (PivotField field in pivotTable.BaseFields)
            {
                if (field.Name.Trim().Equals(nomChamp.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    indexAjoute = pivotTable.AddFieldToArea(type, nomChamp);
                    return true;
                }
            }
            Console.WriteLine($"⚠️ Champ '{nomChamp}' non reconnu par le TCD.");
            return false;
        }
    }
}
