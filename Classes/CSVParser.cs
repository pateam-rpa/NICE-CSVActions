using Direct.Shared;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Direct.CSV.Library.Classes
{
    public static class CSVParser
    {
        /// <summary>
        /// Parses the provided CSV file line by line and imports values to DataTable
        /// </summary>
        /// <param name="filePath">Path of the CSV file</param>
        /// <param name="delimiter">Delimiter by which values in the CSV file are separated</param>
        /// <param name="hasHeader">Informs whether or not provided CSV file contains
        /// header row and if should it be added as a column row in the result DT
        /// </param>
        /// <returns>
        /// DataTable with values from CSV file
        /// </returns>
        /// <exception cref="Exception">Thrown when unexpected error happened while trying
        /// to parse CSV file
        /// </exception>
        public static DataTable ImportCSVToDataTable(string filePath, string delimiter = ",", bool hasHeader = false)
        {
            DataTable dt = new DataTable();
            int lineNumber = 1;

            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new Exception("Error: File Path is empty");
            }

            if (string.IsNullOrEmpty(delimiter))
            {
                delimiter = ",";
            }

            using (StreamReader streamReader = new StreamReader(filePath))
            {
                string currentRawLine;
                bool columnsAdded = false;
                while ((currentRawLine = streamReader.ReadLine()) != null)
                {
                    try
                    {
                        string[] parsedLineValues = SplitCsv(currentRawLine, delimiter);
                        if (columnsAdded)
                        {
                            AddRowToDT(dt, parsedLineValues);
                        }
                        else
                        {
                            if (hasHeader)
                            {
                                AddColumnsToDT(dt, hasHeader, parsedLineValues);
                            }
                            else
                            {
                                AddColumnsToDT(dt, hasHeader, parsedLineValues);
                                AddRowToDT(dt, parsedLineValues);
                            }
                            columnsAdded = true;
                        }
                    }
                    catch (Exception err)
                    {
                        throw new Exception(string.Format("Error: {0}, Line Number: {1}", err.Message, lineNumber.ToString()));
                    }
                    lineNumber++;
                }
                dt.AcceptChanges();
            }
            return dt;
        }

        private static string[] SplitCsv(string line, string delimiter)
        {
            List<string> result = new List<string>();
            StringBuilder currentStr = new StringBuilder("");
            bool inQuotes = false;
            // For each character in the passed line
            for (int i = 0; i < line.Length; i++)
            {
                // check if text contains quotes
                if (line[i] == '\"')
                    inQuotes = !inQuotes;
                else if (line[i] == delimiter[0])
                {
                    // If not in quotes, end of current string, add it to result
                    if (!inQuotes)
                    {
                        result.Add(currentStr.ToString());
                        currentStr.Clear();
                    }
                    else
                        currentStr.Append(line[i]);
                }
                else // Add any other character to current string
                    currentStr.Append(line[i]);
            }
            result.Add(currentStr.ToString());
            return result.ToArray(); // Return array of all strings
        }

        private static void AddColumnsToDT(DataTable dt, bool hasHeader, string[] currentValues)
        {
            for (int columnNumber = 0; columnNumber < currentValues.Length; columnNumber++)
            {
                if (hasHeader)
                {
                    dt.Columns.Add(currentValues[columnNumber].Trim());
                }
                else
                {
                    dt.Columns.Add("Column " + columnNumber.ToString());
                }
            }
        }

        private static void AddRowToDT(DataTable dt, string[] currentValues)
        {
            DataRow row = dt.NewRow();
            for (int columnNumber = 0; columnNumber < currentValues.Length; columnNumber++)
            {
                row[columnNumber] = currentValues[columnNumber];
            }
            dt.Rows.Add(row);
        }

        /// <summary>
        /// Parses the provided CSV file line by line and imports values to List Of Rows
        /// </summary>
        /// <param name="filePath">Path of the CSV file</param>
        /// <param name="delimiter">Delimiter by which values in the CSV file are separated</param>
        /// <param name="skipHeader">Should first row conatain column names or not</param>
        /// <returns>Returns List of Rows with parsed values from CSV
        /// <exception cref="Exception">hrown when unexpected error happened while trying
        /// to parse CSV file
        /// </exception>
        public static DirectCollection<DirectRow> ImportCSVToListOfRows(string filePath, string delimiter = ",", bool skipHeader = false)
        {
            int lineNumber = 1;
            DirectCollection<DirectRow> listOfRows = new DirectCollection<DirectRow>();

            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new Exception("Error: File Path is empty");
            }

            if (string.IsNullOrEmpty(delimiter))
            {
                delimiter = ",";
            }

            using (StreamReader streamReader = new StreamReader(filePath))
            {
                string currentRawLine;
                bool isHeaderSkipped = false;
                while ((currentRawLine = streamReader.ReadLine()) != null)
                {
                    try
                    {
                        string[] parsedLineValues = SplitCsv(currentRawLine, delimiter);
                        DirectRow directRow = new DirectRow();
                        if (skipHeader && !isHeaderSkipped)
                        {
                            isHeaderSkipped = true;
                        }
                        else
                        {
                            directRow.Cells.AddRange(parsedLineValues);
                            listOfRows.Add(directRow);
                        }
                    }
                    catch (Exception err)
                    {
                        throw new Exception(string.Format("Error: {0}, Line Number: {1}", err.Message, lineNumber.ToString()));
                    }
                    lineNumber++;
                }
            }
            return listOfRows;
        }

        /// <summary>
        /// Exports a DirectCollection&lt;DirectRow&gt; to a CSV file.
        /// - Overwrites the file.
        /// - Replaces occurrences of <paramref name="delimiter"/> inside cell values with <paramref name="delimiterReplacer"/>.
        /// - Properly escapes quotes and wraps fields in quotes when they contain quotes or CR/LF.
        /// - Optionally writes a header line.
        /// </summary>
        /// <param name="rows">Rows to export (assumed non-null, with non-null Cells).</param>
        /// <param name="filePath">Destination CSV path.</param>
        /// <param name="delimiter">Field delimiter (default ",").</param>
        /// <param name="delimiterReplacer">Replacement text for delimiter occurrences inside cells.</param>
        /// <param name="header">Optional header row as a plain string (already delimited). Pass null/empty to skip.</param>
        public static void ExportListOfRowsToCsv(
            DirectCollection<DirectRow> rows,
            string filePath,
            string delimiter = ",",
            string delimiterReplacer = ".",
            string header = null)
        {
            const string methodName = nameof(ExportListOfRowsToCsv);

            Loggers.LogDebug(methodName, $"Start. filePath='{filePath}', delimiter='{delimiter}', replacer='{delimiterReplacer}', headerProvided={(string.IsNullOrEmpty(header) ? "no" : "yes")}");

            if (rows == null)
            {
                throw new ArgumentNullException("rows");
            }

            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("File path is empty.", "filePath");
            }

            if (string.IsNullOrEmpty(delimiter))
            {
                delimiter = ",";
            }

            if (string.IsNullOrEmpty(delimiterReplacer))
            {
                delimiter = ".";
            }

            var encoding = Encoding.UTF8;

            try
            {
                var fullPath = Path.GetFullPath(filePath);
                var dir = Path.GetDirectoryName(fullPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                    Loggers.LogDebug(methodName, $"Created directory '{dir}'.");
                }

                using (var writer = new StreamWriter(fullPath, false /* overwrite */, encoding))
                {
                    // Header
                    if (!string.IsNullOrEmpty(header))
                    {
                        writer.WriteLine(SanitizeCell(header, delimiter, delimiterReplacer, true));
                        Loggers.LogDebug(methodName, "Header written.");
                    }

                    int rowIndex = 0;
                    foreach (var row in rows)
                    {
                        // assuming row & Cells are non-null per your choice
                        var sanitized = row.Cells.Select(c => SanitizeCell(c, delimiter, delimiterReplacer, true));
                        writer.WriteLine(string.Join(delimiter, sanitized));

                        // Log every N rows to avoid spam (tweak N as you like)
                        if ((++rowIndex % 500) == 0)
                            Loggers.LogDebug(methodName, $"Written {rowIndex} rows...");
                    }

                    Loggers.LogDebug(methodName, $"Completed. Total rows written: {rowIndex}.");
                }
            }
            catch (Exception ex)
            {
                Loggers.LogError(methodName, ex.Message);
                throw;
            }
        }

        private static string SanitizeCell(string value,
                                           string delimiter,
                                           string delimiterReplacer,
                                           bool quoteIfNeeded)
        {
            if (value == null) value = string.Empty;

            // Replace delimiter occurrences
            if (!string.IsNullOrEmpty(delimiter))
                value = value.Replace(delimiter, delimiterReplacer);

            // Escape quotes by doubling them
            var containsQuote = value.IndexOf('"') >= 0;
            if (containsQuote)
                value = value.Replace("\"", "\"\"");

            // Check for CR/LF
            var containsNewLine = value.IndexOf('\r') >= 0 || value.IndexOf('\n') >= 0;

            // Wrap in quotes if needed
            if (quoteIfNeeded && (containsQuote || containsNewLine))
                return "\"" + value + "\"";

            return value;
        }
    }


}
