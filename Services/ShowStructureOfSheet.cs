using ClosedXML.Excel;

public static class ComptaUtils
{
    public static void AfficherStructureGenerale(XLWorkbook workbook, string sheetName)
{
    var worksheet = workbook.Worksheet(sheetName);
    if (worksheet == null)
    {
        Console.WriteLine($"❌ Feuille '{sheetName}' introuvable.");
        return;
    }

    int maxRowsToAnalyze = 10;
    int lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;

    int titreRowIndex = -1;
    int sousTitreRowIndex = -1;

    // Étape 1 : trouver la ligne la plus "textuelle" (probable ligne de titres)
    for (int rowIndex = 1; rowIndex <= maxRowsToAnalyze; rowIndex++)
    {
        var row = worksheet.Row(rowIndex);
        int textCellCount = 0;
        int numericOrDateCount = 0;

        for (int col = 1; col <= lastColumn; col++)
        {
            var val = row.Cell(col).Value;
            string cellString = row.Cell(col).GetString();
            if (!string.IsNullOrWhiteSpace(cellString))
                textCellCount++;
            else if (double.TryParse(val.ToString(), out _) || DateTime.TryParse(val.ToString(), out _))
                numericOrDateCount++;
        }

        if (textCellCount > 2 && numericOrDateCount == 0) // heuristique simple
        {
            titreRowIndex = rowIndex;
            // Vérifie si la ligne suivante pourrait contenir des sous-titres
            var nextRow = worksheet.Row(rowIndex + 1);
            int subTitleCount = 0;
            for (int col = 1; col <= lastColumn; col++)
            {
                var val = nextRow.Cell(col).GetString().Trim();
                if (!string.IsNullOrEmpty(val))
                    subTitleCount++;
            }

            if (subTitleCount > 0)
                sousTitreRowIndex = rowIndex + 1;

            break;
        }
    }

    if (titreRowIndex == -1)
    {
        Console.WriteLine("❗ Impossible de détecter une ligne de titres.");
        return;
    }

    var mainRow = worksheet.Row(titreRowIndex);
    var subRow = sousTitreRowIndex != -1 ? worksheet.Row(sousTitreRowIndex) : null;

    string? currentMainTitle = null;
    var groupedColumns = new Dictionary<string, List<string>>();

    for (int col = 1; col <= lastColumn; col++)
    {
        string main = mainRow.Cell(col).GetString().Trim();
        string sub = subRow?.Cell(col).GetString().Trim() ?? "";

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