using Direct.CSV.Library.Classes;
using Direct.Shared;
using Direct.Shared.DataTableLibrary;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Direct.CSV.Library
{
    [DirectDom("CSV Actions", "General", false)]
    [ParameterType(false)]
    public static class MainCSV
    {
        [DirectDom("Import CSV To DataTable")]
        [DirectDomMethod("Import CSV {File Path} delimitied with {Delimiter} and containing header {Has Header} into DataTable")]
        [MethodDescription("Imports CSV into DataTable")]
        public static DirectDataTable ImportCSVToDirectDataTable(string filePath, string delimiter, bool hasHeader)
        {
            string methodName = MethodBase.GetCurrentMethod().Name;
            try
            {
                Loggers.LogDebug(methodName, "Attempting to import file: " + filePath + " to DataTable");
                DataTable dt = CSVParser.ImportCSVToDataTable(filePath, delimiter, hasHeader);
                return new DirectDataTable(dt);
            }
            catch (Exception err)
            {
                Loggers.LogError(methodName, "Failed to import. " + err.Message);
                return new DirectDataTable();
            }
        }

        [DirectDom("Import CSV To List Of Rows")]
        [DirectDomMethod("Import CSV {Full File Path} delimitied with {Delimiter} into List Of Rows. Header line should be skipped {Skip Header}")]
        [MethodDescription("Imports CSV into List Of Rows")]
        public static DirectCollection<DirectRow> ImportCSVToListOfRows(string filePath, string delimiter, bool skipHeader)
        {
            string methodName = MethodBase.GetCurrentMethod().Name;
            try
            {
                Loggers.LogDebug(methodName, "Attempting to import file: " + filePath + " to List Of Rows");
                DirectCollection<DirectRow> rows = CSVParser.ImportCSVToListOfRows(filePath, delimiter, skipHeader);
                Loggers.LogDebug(methodName, "Finished importing");
                return rows;
            }
            catch (Exception err)
            {
                Loggers.LogError(methodName, "Failed to import. " + err.Message);
                return new DirectCollection<DirectRow>();
            }
        }

        [DirectDom("Export List Of Rows To CSV")]
        [DirectDomMethod("Export List Of Rows {rows} to CSV at {filePath} using delimiter {delimiter}. Replace delimiter with {delimiterReplacer}. Optional header text {header}")]
        [MethodDescription("Exports the provided List Of Rows to a CSV file")]
        public static bool ExportListOfRowsToCSV(
            DirectCollection<DirectRow> rows,
            string filePath,
            string delimiter,
            string delimiterReplacer,
            string header)
        {
            string methodName = MethodBase.GetCurrentMethod().Name;

            try
            {
                Loggers.LogDebug(methodName, "Wrapper called.");

                CSVParser.ExportListOfRowsToCsv(rows, filePath, delimiter, delimiterReplacer, header);

                Loggers.LogDebug(methodName, "Wrapper finished successfully.");
                return true;
            }
            catch (Exception err)
            {
                Loggers.LogError(methodName, "Failed to export. " + err.Message);
                return false;
            }
        }

        [DirectDom("Convert CSV to Xlsx")]
        [DirectDomMethod("Convert CSV {Full File Path} with delimiter {Delimiter} and save to Excel Spreadsheet {Excel Full File Path}")]
        [MethodDescription("Converts CSV to Xlsx")]
        public static bool ConvertCSVtoXlsx(string filePath, string delimiter, string outputPath)
        {
            string methodName = MethodBase.GetCurrentMethod().Name;
            if (File.Exists(outputPath))
            {
                Loggers.LogDebug(methodName, "Provided excel file already exists");
                return false;
            }

            try
            {
                Loggers.LogDebug(methodName, "Attempting to convert: " + filePath);
                DataTable dt = CSVParser.ImportCSVToDataTable(filePath, delimiter, true);

                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    SheetData sheetData = new SheetData();
                    worksheetPart.Worksheet = new Worksheet(sheetData);

                    Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

                    Sheet sheet = new Sheet()
                    {
                        Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Sheet1"
                    };

                    Row headerRow = new Row();

                    foreach (DataColumn dataColumn in dt.Columns)
                    {
                        Cell headerCell = new Cell()
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(dataColumn.ColumnName)
                        };
                        headerRow.Append(headerCell);
                    }

                    sheetData.Append(headerRow);

                    foreach (string[] dataRow in dt.Select().Select(dr => dr.ItemArray.Select(x => x.ToString()).ToArray()))
                    {
                        Row valuesRow = new Row();
                        foreach (string value in dataRow)
                        {
                            Cell valuesCell = new Cell()
                            {
                                DataType = CellValues.String,
                                CellValue = new CellValue(value)
                            };
                            valuesRow.Append(valuesCell);
                        }
                        sheetData.Append(valuesRow);
                    }
                    sheets.Append(sheet);

                    workbookPart.Workbook.Save();
                    spreadsheetDocument.Save();
                    spreadsheetDocument.Close();
                }
                return true;
            }
            catch (Exception err)
            {
                Loggers.LogError(methodName, "Failed to convert. Error: " + err.Message);
                return false;
            }
        }
    }
}
