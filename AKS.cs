using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml; // EPPlus Namespace
// Alternativ: using ClosedXML.Excel;

namespace ExcelTranslationTool
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Excel Übersetzungstool gestartet");
            
            // Pfade zu den Excel-Dateien
            string inputFilePath = "Eingabedatei.xlsx";
            string translationFilePath = "Uebersetzungstabelle.xlsx";
            string outputFilePath = "Ausgabedatei.xlsx";
            
            try
            {
                // Lizenzeinstellung für EPPlus (bei Verwendung von EPPlus 5+)
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                
                // Hauptfunktionen aufrufen
                var sourceData = ReadSourceExcel(inputFilePath);
                var translationMap = ReadTranslationTable(translationFilePath);
                var translatedData = TranslateData(sourceData, translationMap);
                ExportToExcel(translatedData, outputFilePath);
                
                Console.WriteLine("Übersetzung und Export erfolgreich abgeschlossen!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fehler bei der Verarbeitung: {ex.Message}");
            }
            
            Console.WriteLine("Drücken Sie eine Taste zum Beenden...");
            Console.ReadKey();
        }
        
        /// <summary>
        /// Liest die Quelldaten aus der Excel-Datei ein
        /// </summary>
        static List<Dictionary<string, string>> ReadSourceExcel(string filePath)
        {
            var result = new List<Dictionary<string, string>>();
            
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // Erstes Arbeitsblatt auswählen
                var worksheet = package.Workbook.Worksheets[0];
                
                // Spaltennamen aus der ersten Zeile lesen
                var columnCount = worksheet.Dimension.End.Column;
                var rowCount = worksheet.Dimension.End.Row;
                
                List<string> headers = new List<string>();
                for (int col = 1; col <= columnCount; col++)
                {
                    headers.Add(worksheet.Cells[1, col].Value?.ToString() ?? $"Column{col}");
                }
                
                // Zeilenweise Daten lesen ab Zeile 2 (nach der Kopfzeile)
                for (int row = 2; row <= rowCount; row++)
                {
                    var rowData = new Dictionary<string, string>();
                    
                    for (int col = 1; col <= columnCount; col++)
                    {
                        var value = worksheet.Cells[row, col].Value?.ToString() ?? string.Empty;
                        rowData[headers[col - 1]] = value;
                    }
                    
                    result.Add(rowData);
                }
            }
            
            Console.WriteLine($"{result.Count} Zeilen aus der Quelldatei gelesen.");
            return result;
        }
        
        /// <summary>
        /// Liest die Übersetzungstabelle ein
        /// </summary>
        static Dictionary<string, Dictionary<string, string>> ReadTranslationTable(string filePath)
        {
            var translationMap = new Dictionary<string, Dictionary<string, string>>();
            
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                
                var columnCount = worksheet.Dimension.End.Column;
                var rowCount = worksheet.Dimension.End.Row;
                
                // Spaltennamen aus erster Zeile lesen (ab Spalte 2)
                List<string> targetFields = new List<string>();
                for (int col = 2; col <= columnCount; col++)
                {
                    targetFields.Add(worksheet.Cells[1, col].Value?.ToString() ?? $"Target{col}");
                }
                
                // Übersetzungszuordnungen lesen (ab Zeile 2)
                for (int row = 2; row <= rowCount; row++)
                {
                    var sourceField = worksheet.Cells[row, 1].Value?.ToString();
                    if (string.IsNullOrEmpty(sourceField)) continue;
                    
                    var mappings = new Dictionary<string, string>();
                    
                    for (int col = 2; col <= columnCount; col++)
                    {
                        var targetValue = worksheet.Cells[row, col].Value?.ToString() ?? string.Empty;
                        mappings[targetFields[col - 2]] = targetValue;
                    }
                    
                    translationMap[sourceField] = mappings;
                }
            }
            
            Console.WriteLine($"Übersetzungstabelle mit {translationMap.Count} Einträgen geladen.");
            return translationMap;
        }
        
        /// <summary>
        /// Führt die Übersetzung der Quelldaten anhand der Übersetzungstabelle durch
        /// </summary>
        static List<Dictionary<string, string>> TranslateData(
            List<Dictionary<string, string>> sourceData, 
            Dictionary<string, Dictionary<string, string>> translationMap)
        {
            var result = new List<Dictionary<string, string>>();
            
            foreach (var row in sourceData)
            {
                var translatedRow = new Dictionary<string, string>();
                
                foreach (var sourceKey in row.Keys)
                {
                    var sourceValue = row[sourceKey];
                    
                    // Prüfen, ob für diesen Wert eine Übersetzung existiert
                    if (translationMap.TryGetValue(sourceValue, out var translations))
                    {
                        // Übersetzungen für diesen Wert einfügen
                        foreach (var translation in translations)
                        {
                            translatedRow[translation.Key] = translation.Value;
                        }
                    }
                    
                    // Original-Feld/Wert ebenfalls hinzufügen
                    translatedRow[sourceKey] = sourceValue;
                }
                
                result.Add(translatedRow);
            }
            
            Console.WriteLine($"{result.Count} Zeilen übersetzt.");
            return result;
        }
        
        /// <summary>
        /// Exportiert die übersetzten Daten in eine neue Excel-Datei
        /// </summary>
        static void ExportToExcel(List<Dictionary<string, string>> data, string filePath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Übersetzung");
                
                // Alle eindeutigen Spaltenüberschriften sammeln
                var allHeaders = new HashSet<string>();
                foreach (var row in data)
                {
                    foreach (var key in row.Keys)
                    {
                        allHeaders.Add(key);
                    }
                }
                
                var headersList = new List<string>(allHeaders);
                
                // Überschriften in erste Zeile schreiben
                for (int i = 0; i < headersList.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = headersList[i];
                    // Formatierung der Kopfzeile
                    worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                }
                
                // Daten zeilenweise schreiben
                for (int rowIndex = 0; rowIndex < data.Count; rowIndex++)
                {
                    var rowData = data[rowIndex];
                    
                    for (int colIndex = 0; colIndex < headersList.Count; colIndex++)
                    {
                        var header = headersList[colIndex];
                        if (rowData.TryGetValue(header, out var cellValue))
                        {
                            worksheet.Cells[rowIndex + 2, colIndex + 1].Value = cellValue;
                        }
                    }
                }
                
                // Autofit für bessere Lesbarkeit
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                
                // Datei speichern
                package.SaveAs(new FileInfo(filePath));
            }
            
            Console.WriteLine($"Daten wurden erfolgreich nach {filePath} exportiert.");
        }
    }
}