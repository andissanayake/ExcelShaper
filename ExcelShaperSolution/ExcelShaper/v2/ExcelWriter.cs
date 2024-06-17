using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Reflection;

namespace ExcelShaperLib
{
    public partial class ExcelShaper
    {
        public static void WriteToExcel<T>(string filePath, List<T> data, Dictionary<string, string>? headers = null, Func<T, PropertyInfo, string> complexPropertyConverter = null)
        {
            using SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
            WorkbookPart workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
            Sheet sheet = new() { Id = document.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
            sheets.Append(sheet);

            var columns = WriteDataToWorksheet(worksheetPart, data, headers, complexPropertyConverter);

            AddTableDefinition(worksheetPart, data.Count, columns);

            workbookPart.Workbook.Save();
        }

        private static int WriteDataToWorksheet<T>(WorksheetPart worksheetPart, List<T> data, Dictionary<string, string>? headers, Func<T, PropertyInfo, string> complexPropertyConverter)
        {
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            PropertyInfo[] properties = typeof(T).GetProperties();
            AddHeaderRow(sheetData, properties, headers);

            foreach (T item in data)
            {
                Row row = new();
                foreach (var property in properties)
                {
                    Cell cell = new();
                    object value = property.GetValue(item);
                    string cellValue = (complexPropertyConverter != null && value != null) ? complexPropertyConverter(item, property) : value?.ToString() ?? string.Empty;
                    cell.CellValue = new CellValue(cellValue);
                    cell.DataType = CellValues.String;
                    row.AppendChild(cell);
                }
                sheetData.AppendChild(row);
            }

            return properties.Length;
        }

        private static void AddHeaderRow(SheetData sheetData, PropertyInfo[] properties, Dictionary<string, string>? headers)
        {
            Row headerRow = new();
            foreach (var property in properties)
            {
                Cell cell = new()
                {
                    CellValue = new CellValue(headers != null && headers.ContainsKey(property.Name) ? headers[property.Name] : property.Name),
                    DataType = CellValues.String,
                    StyleIndex = 1
                };
                headerRow.AppendChild(cell);
            }
            sheetData.AppendChild(headerRow);
        }

        private static void AddTableDefinition(WorksheetPart worksheetPart, int rowCount, int columnCount)
        {
            var tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>();
            string tableRange = $"A1:{(char)('A' + columnCount - 1)}{rowCount + 1}";
            Table table = new()
            {
                Id = 1,
                Name = "Table1",
                DisplayName = "Table1",
                Reference = tableRange,
                TotalsRowShown = false
            };

            TableColumns tableColumns = new() { Count = (uint)columnCount };
            for (int i = 0; i < columnCount; i++)
            {
                tableColumns.Append(new TableColumn() { Id = (uint)(i + 1), Name = $"Column{i + 1}" });
            }

            table.AppendChild(tableColumns);
            table.AppendChild(new TableStyleInfo()
            {
                Name = "TableStyleMedium9",
                ShowFirstColumn = false,
                ShowLastColumn = false,
                ShowRowStripes = true,
                ShowColumnStripes = false
            });

            tableDefinitionPart.Table = table;
        }
    }
}
