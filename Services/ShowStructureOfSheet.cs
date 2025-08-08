using Aspose.Cells;
using System;
using System.Collections.Generic;

public static class ComptaUtils
{
    public static void AfficherStructureGenerale(Workbook workbook, string sheetName)
    {
        var worksheet = workbook.Worksheets[sheetName];
        if (worksheet == null)
        {
            Console.WriteLine($"❌ Feuille '{sheetName}' introuvable.");
            return;
        }

        int maxRowsToAnalyze = 10;
        var plage = worksheet.Cells.MaxDisplayRange;
        int lastColumn = plage?.ColumnCount ?? 0;

        int titreRowIndex = -1;
        int sousTitreRowIndex = -1;

        // ➤ Étape 1 : détecter la ligne de titres
        for (int rowIndex = 0; rowIndex < maxRowsToAnalyze; rowIndex++)
        {
            int textCellCount = 0;
            int numericOrDateCount = 0;

            for (int col = 0; col < lastColumn; col++)
            {
                var cell = worksheet.Cells[rowIndex, col];
                string cellString = cell.StringValue.Trim();
                if (!string.IsNullOrWhiteSpace(cellString))
                    textCellCount++;
                else if (cell.Type == CellValueType.IsNumeric || cell.Type == CellValueType.IsDateTime)
                    numericOrDateCount++;
            }

            if (textCellCount > 2 && numericOrDateCount == 0)
            {
                titreRowIndex = rowIndex;

                var nextRow = rowIndex + 1;
                if (nextRow < worksheet.Cells.MaxDataRow)
                {
                    int subTitleCount = 0;
                    for (int col = 0; col < lastColumn; col++)
                    {
                        string val = worksheet.Cells[nextRow, col].StringValue.Trim();
                        if (!string.IsNullOrEmpty(val))
                            subTitleCount++;
                    }

                    if (subTitleCount > 0)
                        sousTitreRowIndex = nextRow;
                }

                break;
            }
        }

        if (titreRowIndex == -1)
        {
            Console.WriteLine("❗ Impossible de détecter une ligne de titres.");
            return;
        }

        var groupedColumns = new Dictionary<string, List<string>>();
        string? currentMainTitle = null;

        for (int col = 0; col < lastColumn; col++)
        {
            string main = worksheet.Cells[titreRowIndex, col].StringValue.Trim();
            string sub = sousTitreRowIndex != -1 ? worksheet.Cells[sousTitreRowIndex, col].StringValue.Trim() : "";

            if (!string.IsNullOrEmpty(main))
            {
                currentMainTitle = main;
                if (!groupedColumns.ContainsKey(currentMainTitle))
                    groupedColumns[currentMainTitle] = new List<string>();
            }

            if (!string.IsNullOrEmpty(sub))
            {
                if (string.IsNullOrEmpty(currentMainTitle))
                    currentMainTitle = "(Sans titre)";
                if (!groupedColumns.ContainsKey(currentMainTitle))
                    groupedColumns[currentMainTitle] = new List<string>();
                groupedColumns[currentMainTitle].Add(sub);
            }
            else if (!string.IsNullOrEmpty(main))
            {
                if (string.IsNullOrEmpty(currentMainTitle))
                    currentMainTitle = "(Sans titre)";
                if (!groupedColumns.ContainsKey(currentMainTitle))
                    groupedColumns[currentMainTitle] = new List<string>();
                if (!groupedColumns[currentMainTitle].Contains("(Aucune sous-colonne)"))
                    groupedColumns[currentMainTitle].Add("(Aucune sous-colonne)");
            }
        }

        Console.WriteLine($"\n📊 Structure détectée automatiquement (feuille '{sheetName}') :\n");
        foreach (var entry in groupedColumns)
        {
            Console.WriteLine($"📁 {entry.Key}");
            foreach (var sub in entry.Value)
            {
                Console.WriteLine($"   └── 📄 {sub}");
            }
        }
    }
}
