using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using System.Linq;

namespace ComptaFlow.Services
{
    public class TCDGeneratorIntegra
    {
        public void GenererTCD(string cheminFichierExcel, string cheminSortie)
        {
            try
            {
                var workbookSource = new Workbook(cheminFichierExcel);
                var workbookFinal = new Workbook();
                workbookFinal.Worksheets.Clear();

                SetDefaultStyle(workbookFinal);

                // Copier toutes les feuilles sauf "Feuil1"
                foreach (Worksheet feuilleSource in workbookSource.Worksheets)
                {
                    if (feuilleSource.Name.Equals("Feuil1", StringComparison.OrdinalIgnoreCase))
                        continue; // Ignorer "Feuil1"

                    string nomFeuillePropre = NettoyerNomFeuille(feuilleSource.Name);
                    var feuilleCopie = workbookFinal.Worksheets.Add(nomFeuillePropre);
                    feuilleCopie.Copy(feuilleSource);
                }

                // Trouver la feuille copiée "Rapport Op_Trait"
                Worksheet feuilleCopieTCD = workbookFinal.Worksheets["Rapport Op_Trait"];
                if (feuilleCopieTCD == null)
                {
                    Console.WriteLine("❌ Feuille 'Rapport Op_Trait' introuvable dans le classeur copié.");
                    return;
                }

                int lastRow = feuilleCopieTCD.Cells.MaxDataRow;
                int lastCol = feuilleCopieTCD.Cells.MaxDataColumn;

                if (lastRow < 0 || lastCol < 0)
                {
                    Console.WriteLine("❌ Plage de données vide dans la feuille 'Rapport Op_Trait'.");
                    return;
                }

                int ligneEntetes = TrouverLigneEntetes(feuilleCopieTCD, lastCol, "Montant");
                if (ligneEntetes == -1)
                {
                    Console.WriteLine("❌ Impossible de trouver la ligne d'en-têtes contenant 'Montant' dans la feuille 'Rapport Op_Trait'.");
                    return;
                }

                string debutPlage = CellsHelper.CellIndexToName(ligneEntetes, 0);
                string finPlage = CellsHelper.CellIndexToName(lastRow, lastCol);
                string plageAdresse = $"'{feuilleCopieTCD.Name}'!{debutPlage}:{finPlage}";

                // Ajouter une nouvelle feuille pour le TCD
                Worksheet feuilleTCD = workbookFinal.Worksheets.Add("TCD_RapportOpTrait");

                // Créer le TCD
                int indexPivot = feuilleTCD.PivotTables.Add(plageAdresse, "A1", "Pivot_RapportOpTrait");
                PivotTable pivotTable = feuilleTCD.PivotTables[indexPivot];

                pivotTable.ShowInCompactForm();

                // Ajout des champs en filtre
                AjouterChampSiExiste(pivotTable, PivotFieldType.Page, "Franchisé");
                AjouterChampSiExiste(pivotTable, PivotFieldType.Page, "Agence");
                AjouterChampSiExiste(pivotTable, PivotFieldType.Page, "Service");

                // Ajout des champs en ligne
                AjouterChampSiExiste(pivotTable, PivotFieldType.Row, "Date");

                // Ajout du champ en valeur
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

                // Totaux généraux visibles
                pivotTable.ShowRowGrandTotals = true;
                pivotTable.ShowColumnGrandTotals = true;

                pivotTable.RefreshData();
                pivotTable.CalculateData();

                workbookFinal.Save(cheminSortie);

                // Supprimer la dernière feuille avec ClosedXML
                SupprimerDerniereFeuilleAvecClosedXml(cheminSortie);

                Console.WriteLine($"✅ TCD généré avec succès : {cheminSortie}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Une erreur est survenue lors de la génération du fichier.");
                Console.WriteLine($"Détails : {ex.Message}");
            }
        }

        private void SupprimerDerniereFeuilleAvecClosedXml(string cheminFichier)
        {
            using var workbook = new XLWorkbook(cheminFichier);

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

        private string NettoyerNomFeuille(string nom)
        {
            string nettoye = Regex.Replace(nom, @"[\\\/\*\[\]\?:']", "_");
            if (nettoye.Length > 31)
                nettoye = nettoye.Substring(0, 31);
            return nettoye;
        }

        private int TrouverLigneEntetes(Worksheet feuille, int maxCol, string champRecherche)
        {
            int maxRowToCheck = feuille.Cells.MaxDataRow;
            for (int row = 0; row <= maxRowToCheck; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    var cellValue = feuille.Cells[row, col].StringValue?.Trim();
                    if (!string.IsNullOrEmpty(cellValue) &&
                        cellValue.Equals(champRecherche, StringComparison.OrdinalIgnoreCase))
                    {
                        return row;
                    }
                }
            }
            return -1;
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
