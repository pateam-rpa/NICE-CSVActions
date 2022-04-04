﻿using Direct.Shared;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
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
    }
}
