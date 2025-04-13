using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using OfficeOpenXml; // EPPlus Namespace

namespace ExcelTranslationTool
{
    class Program
    {
        // Konfiguration für Regex-Muster
        static readonly Dictionary<string, string> ValidationPatterns = new Dictionary<string, string>
        {
            // Beispiel-Validierungsregeln:
            { "Name", @"^[A-Za-z\s-]{2,50}$" },                     // Namen: nur Buchstaben, Leerzeichen, Bindestrich, 2-50 Zeichen
            { "Email", @"^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$" },       // Standard E-Mail-Format
            { "Telefon", @"^\+?[0-9\s-]{6,20}$" },                  // Telefonnummern
            { "Postleitzahl", @"^\d{4,5}$" },                       // 4-5 stellige Postleitzahlen
            { "Datum", @"^\d{1,2}\.\d{1,2}\.\d{4}$" }               // Deutsches Datumsformat (TT.MM.JJJJ)
            // Füge hier weitere Validierungsmuster nach Bedarf hinzu
        };
        
        // Konsole-Farbkonfiguration für verschiedene Nachrichten
        static readonly ConsoleColor ErrorColor = ConsoleColor.Red;
        static readonly ConsoleColor WarningColor = ConsoleColor.Yellow;
        static readonly ConsoleColor SuccessColor = ConsoleColor.Green;
        static readonly ConsoleColor InfoColor = ConsoleColor.Cyan;
        
        static void Main(string[] args)
        {
            Console.WriteLine("Excel Übersetzungstool gestartet");
            
            // Pfade zu den Excel-Dateien
            string inputFilePath = "Eingabedatei.xlsx";
            string translationFilePath = "Uebersetzungstabelle.xlsx";
            string outputFilePath = "Ausgabedatei.xlsx";
            
            // Optionale Log-Datei für Validierungsfehler
            string logFilePath = "validation_errors.log";
            List<string> validationErrors = new List<string>();
            
            try
            {
                // Lizenzeinstellung für EPPlus (bei Verwendung von EPPlus 5+)
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                
                // Hauptfunktionen aufrufen
                PrintColorMessage("Beginne Einlesen der Quelldaten...", InfoColor);
                var sourceData = ReadSourceExcel(inputFilePath, validationErrors);
                
                PrintColorMessage("Beginne Einlesen der Übersetzungstabelle...", InfoColor);
                var translationMap = ReadTranslationTable(translationFilePath, validationErrors);
                
                PrintColorMessage("Beginne Übersetzung der Daten...", InfoColor);
                var translatedData = TranslateData(sourceData, translationMap);
                
                PrintColorMessage("Exportiere übersetzte Daten...", InfoColor);
                ExportToExcel(translatedData, outputFilePath);
                
                // Validierungsergebnisse ausgeben
                if (validationErrors.Count > 0)
                {
                    PrintColorMessage($"Es wurden {validationErrors.Count} Validierungsprobleme gefunden. Details werden in {logFilePath} gespeichert.", WarningColor);
                    File.WriteAllLines(logFilePath, validationErrors);
                }
                else
                {
                    PrintColorMessage("Alle Daten haben die Validierung bestanden.", SuccessColor);
                }
                
                PrintColorMessage("Übersetzung und Export erfolgreich abgeschlossen!", SuccessColor);
            }
            catch (Exception ex)
            {
                PrintColorMessage($"Kritischer Fehler bei der Verarbeitung: {ex.Message}", ErrorColor);
                Console.WriteLine(ex.StackTrace);
            }
            
            Console.WriteLine("\nDrücken Sie eine Taste zum Beenden...");
            Console.ReadKey();
        }
        
        /// <summary>
        /// Gibt eine farbige Nachricht in der Konsole aus
        /// </summary>
        static void PrintColorMessage(string message, ConsoleColor color)
        {
            var originalColor = Console.ForegroundColor;
            Console.ForegroundColor = color;
            Console.WriteLine(message);
            Console.ForegroundColor = originalColor;
        }
        
        /// <summary>
        /// Liest die Quelldaten aus der Excel-Datei ein und validiert sie
        /// </summary>
        static List<Dictionary<string, string>> ReadSourceExcel(string filePath, List<string> validationErrors)
        {
            var result = new List<Dictionary<string, string>>();
            
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Die Quelldatei wurde nicht gefunden: {filePath}");
            }
            
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
                    bool rowHasValidationError = false;
                    
                    for (int col = 1; col <= columnCount; col++)
                    {
                        var columnName = headers[col - 1];
                        var value = worksheet.Cells[row, col].Value?.ToString() ?? string.Empty;
                        rowData[columnName] = value;
                        
                        // Regex-Validierung durchführen, wenn Muster für diese Spalte definiert ist
                        if (ValidationPatterns.TryGetValue(columnName, out string pattern) && !string.IsNullOrEmpty(value))
                        {
                            if (!Regex.IsMatch(value, pattern))
                            {
                                string errorMessage = $"Zeile {row}, Spalte '{columnName}': Wert '{value}' entspricht nicht dem erwarteten Format.";
                                validationErrors.Add(errorMessage);
                                rowHasValidationError = true;
                                
                                // Ausgabe des Fehlers in der Konsole
                                PrintColorMessage(errorMessage, WarningColor);
                            }
                        }
                        
                        // Spezialfall-Validierungen (Beispiele)
                        if (columnName == "Kategorie" && value == "LEER")
                        {
                            string message = $"Zeile {row}, Spalte 'Kategorie': Wert 'LEER' gefunden - Sonderbehandlung erforderlich.";
                            validationErrors.Add(message);
                            PrintColorMessage(message, WarningColor);
                        }
                        
                        if (columnName == "Preis" && !string.IsNullOrEmpty(value))
                        {
                            if (!decimal.TryParse(value.Replace(",", "."), out _))
                            {
                                string errorMessage = $"Zeile {row}, Spalte 'Preis': Wert '{value}' ist keine gültige Zahl.";
                                validationErrors.Add(errorMessage);
                                rowHasValidationError = true;
                                PrintColorMessage(errorMessage, WarningColor);
                            }
                        }
                    }
                    
                    // Zusätzliche Validierung für Kombinationen von Feldern
                    if (rowData.TryGetValue("StartDatum", out string startDatum) && 
                        rowData.TryGetValue("EndDatum", out string endDatum))
                    {
                        if (!string.IsNullOrEmpty(startDatum) && !string.IsNullOrEmpty(endDatum))
                        {
                            if (DateTime.TryParse(startDatum, out DateTime start) && 
                                DateTime.TryParse(endDatum, out DateTime end))
                            {
                                if (start > end)
                                {
                                    string errorMessage = $"Zeile {row}: StartDatum '{startDatum}' liegt nach EndDatum '{endDatum}'.";
                                    validationErrors.Add(errorMessage);
                                    rowHasValidationError = true;
                                    PrintColorMessage(errorMessage, WarningColor);
                                }
                            }
                        }
                    }
                    
                    if (rowHasValidationError)
                    {
                        // Optional: Markiere Zeilen mit Validierungsfehlern in den Daten
                        rowData["_HasValidationErrors"] = "true";
                    }
                    
                    result.Add(rowData);
                }
            }
            
            Console.WriteLine($"{result.Count} Zeilen aus der Quelldatei gelesen.");
            return result;
        }
        
        /// <summary>
        /// Liest die Übersetzungstabelle ein und validiert ihre Struktur
        /// </summary>
        static Dictionary<string, Dictionary<string, string>> ReadTranslationTable(string filePath, List<string> validationErrors)
        {
            var translationMap = new Dictionary<string, Dictionary<string, string>>();
            
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Die Übersetzungstabelle wurde nicht gefunden: {filePath}");
            }
            
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                
                var columnCount = worksheet.Dimension.End.Column;
                var rowCount = worksheet.Dimension.End.Row;
                
                if (columnCount < 2)
                {
                    throw new InvalidDataException("Die Übersetzungstabelle muss mindestens 2 Spalten haben.");
                }
                
                // Spaltennamen aus erster Zeile lesen (ab Spalte 2)
                List<string> targetFields = new List<string>();
                for (int col = 2; col <= columnCount; col++)
                {
                    targetFields.Add(worksheet.Cells[1, col].Value?.ToString() ?? $"Target{col}");
                }
                
                // Regex-Muster für die Validierung von Quellwerten
                string sourceValuePattern = @"^[A-Za-z0-9\s.\-_]{1,100}$"; // Beispiel: Buchstaben, Zahlen, Leerzeichen, Sonderzeichen
                
                // Übersetzungszuordnungen lesen (ab Zeile 2)
                for (int row = 2; row <= rowCount; row++)
                {
                    var sourceField = worksheet.Cells[row, 1].Value?.ToString();
                    
                    // Überprüfung, ob der Quellwert leer ist
                    if (string.IsNullOrEmpty(sourceField))
                    {
                        string message = $"Übersetzungstabelle Zeile {row}: Quellwert ist leer, diese Zeile wird übersprungen.";
                        validationErrors.Add(message);
                        PrintColorMessage(message, WarningColor);
                        continue;
                    }
                    
                    // Validierung des Quellwerts mit Regex
                    if (!Regex.IsMatch(sourceField, sourceValuePattern))
                    {
                        string message = $"Übersetzungstabelle Zeile {row}: Quellwert '{sourceField}' enthält ungültige Zeichen.";
                        validationErrors.Add(message);
                        PrintColorMessage(message, WarningColor);
                    }
                    
                    // Überprüfung auf Duplikate
                    if (translationMap.ContainsKey(sourceField))
                    {
                        string message = $"Übersetzungstabelle Zeile {row}: Doppelter Quellwert '{sourceField}' gefunden. Vorherige Definition wird überschrieben.";
                        validationErrors.Add(message);
                        PrintColorMessage(message, WarningColor);
                    }
                    
                    var mappings = new Dictionary<string, string>();
                    bool hasAtLeastOneMapping = false;
                    
                    for (int col = 2; col <= columnCount; col++)
                    {
                        var targetValue = worksheet.Cells[row, col].Value?.ToString() ?? string.Empty;
                        
                        // Überprüfung, ob das Zielfeld nicht leer ist
                        if (!string.IsNullOrEmpty(targetValue))
                        {
                            hasAtLeastOneMapping = true;
                            
                            // Spaltenspezifische Regex-Validierung
                            string columnName = targetFields[col - 2];
                            if (ValidationPatterns.TryGetValue(columnName, out string pattern))
                            {
                                if (!Regex.IsMatch(targetValue, pattern))
                                {
                                    string message = $"Übersetzungstabelle Zeile {row}, Spalte '{columnName}': Wert '{targetValue}' entspricht nicht dem erwarteten Format.";
                                    validationErrors.Add(message);
                                    PrintColorMessage(message, WarningColor);
                                }
                            }
                        }
                        
                        mappings[targetFields[col - 2]] = targetValue;
                    }
                    
                    // Warnung, wenn keine Zielwerte vorhanden sind
                    if (!hasAtLeastOneMapping)
                    {
                        string message = $"Übersetzungstabelle Zeile {row}: Quellwert '{sourceField}' hat keine Übersetzungen.";
                        validationErrors.Add(message);
                        PrintColorMessage(message, WarningColor);
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
            int missingTranslationsCount = 0;
            
            foreach (var row in sourceData)
            {
                var translatedRow = new Dictionary<string, string>();
                bool rowHasMissingTranslation = false;
                
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
                    else if (!string.IsNullOrEmpty(sourceValue) && sourceValue != "0" && !sourceKey.StartsWith("_"))
                    {
                        // Fehlende Übersetzung für nicht-leere Werte
                        rowHasMissingTranslation = true;
                        PrintColorMessage($"Keine Übersetzung gefunden für '{sourceValue}' in Feld '{sourceKey}'", WarningColor);
                    }
                    
                    // Original-Feld/Wert ebenfalls hinzufügen
                    translatedRow[sourceKey] = sourceValue;
                }
                
                if (rowHasMissingTranslation)
                {
                    missingTranslationsCount++;
                    translatedRow["_MissingTranslation"]