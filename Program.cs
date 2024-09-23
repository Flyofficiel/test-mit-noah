using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        // Der absolute Pfad zur Excel-Datei
        var filePath = @"/Users/Lucas/Seafile/CorporateAi/Receipts/DATEV/Buchungen/TA2021.xlsx"; 

        // Überprüfen, ob die Datei existiert
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Datei nicht gefunden: {filePath}");
            return;
        }

        // Limitierung (z.B. nur die ersten 10 Zeilen und 5 Spalten auslesen)
        int maxRows = 10;
        int maxCols = 23;

        // Datei mit EPPlus laden und verarbeiten
        FileInfo fileInfo = new FileInfo(filePath);
        using (ExcelPackage package = new ExcelPackage(fileInfo))
        {
            // Lade das erste Arbeitsblatt (Sheet)
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            // Bestimme die Anzahl der Zeilen und Spalten im Arbeitsblatt
            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;

            // Wende die Limitierung auf Zeilen und Spalten an
            int rowsToRead = Math.Min(rowCount, maxRows);
            int colsToRead = Math.Min(colCount, maxCols);

            // Speicherstruktur für die Variablen
            var rowDataList = new List<Dictionary<string, string>>();

            // Iteriere über alle Zeilen und Spalten innerhalb der Limitierung
            for (int row = 2; row <= rowsToRead; row++)
            {
                // Dictionary zur Speicherung der Zeilenwerte als Schlüssel/Wert-Paar
                var rowData = new Dictionary<string, string>();

                for (int col = 1; col <= colsToRead; col++)
                {
                    // Lese den Inhalt der Zelle (row, col)
                    var cellValue = worksheet.Cells[row, col].Text;
                    string columnName = worksheet.Cells[1, col].Text; // Spaltennamen aus erster Zeile

                    // Speichere den Wert mit dem Spaltennamen als Schlüssel
                    rowData[columnName] = cellValue;
                }

                // Füge die Zeile der Liste hinzu
                rowDataList.Add(rowData);
            }

            // Zugriff auf Variablen (Beispiel für die erste Zeile)
            if (rowDataList.Count > 0)
            {
                var firstRow = rowDataList[0];
                if (firstRow.ContainsKey("Buchungsdatum")) // Ersetze "Buchungsdatum" durch den tatsächlichen Spaltennamen
                {
                    string buchungsdatum = firstRow["Buchungsdatum"];
                    Console.WriteLine($"Buchungsdatum der ersten Zeile: {buchungsdatum}");
                }
            }
        }
    }
}
