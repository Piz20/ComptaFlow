using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.Text.RegularExpressions;
using System.Drawing;

namespace ComptaFlow.Services
{
    public class TCDGeneratorSage
    {
        public void GenererTCDAvecFeuilPrecedente(string cheminFichierExcel, string cheminSortie)
        {
            var workbookSource = new Workbook(cheminFichierExcel);
            var workbookFinal = new Workbook();
            workbookFinal.Worksheets.Clear();

            // ➤ Définir le style par défaut pour tout le workbook (Calibri 11)
            Style styleDefaut = workbookFinal.CreateStyle();
            styleDefaut.Font.Name = "Calibri";
            styleDefaut.Font.Size = 11;
            workbookFinal.DefaultStyle = styleDefaut;

            int compteurFeuil = 1;

            foreach (Worksheet feuilleSource in workbookSource.Worksheets)
            {
                if (feuilleSource.Name.StartsWith("Feuil"))
                    continue;

                NettoyerEtRenommerFeuille(feuilleSource, ref compteurFeuil);

                Worksheet copieFeuille = workbookFinal.Worksheets.Add(feuilleSource.Name);
                copieFeuille.Copy(feuilleSource);

                // ➤ Appliquer Calibri 11 à toute la feuille copiée
                AppliquerStyleCalibriFeuille(copieFeuille);

                string nomFeuilTCD = $"Feuil{compteurFeuil++}";
                Worksheet feuilleTCD = workbookFinal.Worksheets.Add(nomFeuilTCD);

                var plageDonnees = feuilleSource.Cells.MaxDisplayRange;
                if (plageDonnees == null || plageDonnees.RowCount == 0 || plageDonnees.ColumnCount == 0)
                    continue;

                string plageAdresse = $"'{feuilleSource.Name}'!{plageDonnees.Address}";

                try
                {
                    // ➤ Ajouter le TCD
                    int indexPivot = feuilleTCD.PivotTables.Add(plageAdresse, "A1", "PivotTable1");
                    PivotTable pivotTable = feuilleTCD.PivotTables[indexPivot];

                    // ➤ Mode compact activé (appel de méthode)
                    pivotTable.ShowInCompactForm();

                    // ➤ Ajout des champs "Étiquette de lignes" : date en 1er, libellé en 2e
                    pivotTable.AddFieldToArea(PivotFieldType.Row, "Date");
                    pivotTable.AddFieldToArea(PivotFieldType.Row, "Libellé écriture");

                    // ➤ Trier automatiquement les étiquettes de lignes de A à Z
                    foreach (PivotField rowField in pivotTable.RowFields)
                    {
                        rowField.IsAutoSort = true;
                        rowField.IsAscendSort = true; // Tri croissant (A à Z)
                    }

                    // ➤ Ajouter champ valeur : Somme de "Montant signé (XAF)"
                    int dataFieldIndex = pivotTable.AddFieldToArea(PivotFieldType.Data, "Montant signé (XAF)");
                    pivotTable.DataFields[dataFieldIndex].Function = ConsolidationFunction.Sum;

                    // ➤ Renommer l'étiquette du champ en français
                    pivotTable.DataFields[dataFieldIndex].DisplayName = "Somme de Montant signé (XAF)";

                    // Afficher les sous-totaux
                    // ➤ Afficher les sous-totaux uniquement pour le champ "Date"
                    foreach (PivotField rowField in pivotTable.RowFields)
                    {
                        // C'est le champ "Date"
                        if (rowField.Name == "Date")
                        {
                            rowField.SetSubtotals(PivotFieldSubtotalType.Sum, true); // Active les sous-totaux de type Somme
                            rowField.ShowSubtotalAtTop = true; // Affiche les sous-totaux en haut (ou en bas, si vous préférez)
                        }
                        // C'est le champ "Libellé écriture"
                        else if (rowField.Name == "Libellé écriture")
                        {
                            rowField.SetSubtotals(PivotFieldSubtotalType.None, true); // Désactive les sous-totaux
                        }
                    }

                    // ➤ Garder le grand total général
                    pivotTable.ShowRowGrandTotals = true;
                    pivotTable.ShowColumnGrandTotals = true;

                    // ➤ Garder uniquement le total général
                    pivotTable.ShowRowGrandTotals = true;
                    pivotTable.ShowColumnGrandTotals = true;

                    // ➤ Rafraîchir et calculer
                    pivotTable.RefreshData();
                    pivotTable.CalculateData();

                    // ➤ Appliquer Calibri 11 à toute la feuille TCD
                    AppliquerStyleCalibriFeuille(feuilleTCD);

                    int startRow = pivotTable.TableRange2.StartRow;
                    int endRow = pivotTable.TableRange2.EndRow;
                    int maxCol = feuilleTCD.Cells.MaxColumn;

                    int colDate = 0; // première colonne (Date)
                    int colLibelle = 1; // deuxième colonne (Libellé écriture)
                    int colMontant = 2; // la colonne de données (à ajuster si différent)

                    // ➤ Style gras avec Calibri 11
                    Style styleGras = feuilleTCD.Workbook.CreateStyle();
                    styleGras.Font.IsBold = true;
                    styleGras.Font.Name = "Calibri";
                    styleGras.Font.Size = 11;
                    StyleFlag flagGras = new StyleFlag() { FontBold = true, FontName = true, FontSize = true };

                    for (int row = startRow; row <= endRow; row++)
                    {
                        Cell cellDate = feuilleTCD.Cells[row, colDate];
                        Cell cellLibelle = feuilleTCD.Cells[row, colLibelle];
                        Cell cellMontant = feuilleTCD.Cells[row, colMontant];

                        // ➤ Si colonne Date vide ou fusionnée + colonne Libellé vide ou fusionnée + colonne Montant non vide => probablement une ligne de total
                        bool isDateVide = cellDate.IsMerged || string.IsNullOrWhiteSpace(cellDate.StringValue);
                        bool isLibelleVide = cellLibelle.IsMerged || string.IsNullOrWhiteSpace(cellLibelle.StringValue);
                        bool hasMontant = cellMontant.Type == CellValueType.IsNumeric && cellMontant.DoubleValue != 0;

                        if (isDateVide && isLibelleVide && hasMontant)
                        {
                            // ➤ Appliquer le style gras à toute la ligne
                            for (int col = 0; col <= maxCol; col++)
                            {
                                Cell cell = feuilleTCD.Cells[row, col];
                                cell.SetStyle(styleGras, flagGras);
                            }
                        }
                    }

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

            workbookFinal.Save(cheminSortie);
            Console.WriteLine($"Fichier généré : {cheminSortie}");
        }

        private void AppliquerStyleCalibriFeuille(Worksheet feuille)
        {
            // ➤ Style Calibri 11 pour toute la feuille
            Style styleCalibriFeuille = feuille.Workbook.CreateStyle();
            styleCalibriFeuille.Font.Name = "Calibri";
            styleCalibriFeuille.Font.Size = 11;

            StyleFlag flag = new StyleFlag();
            flag.FontName = true;
            flag.FontSize = true;

            // ➤ Appliquer à toutes les cellules utilisées
            if (feuille.Cells.MaxDisplayRange != null)
            {
                feuille.Cells.ApplyStyle(styleCalibriFeuille, flag);
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