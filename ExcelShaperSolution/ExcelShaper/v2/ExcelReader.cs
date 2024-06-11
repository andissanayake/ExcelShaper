using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelShaper
{
    public class Engine2
    {
        public static List<List<string>> ReadExcelFileByIndex(string filePath, int sheetIndex = 0)
        {
            List<List<string>> excelData = new();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorksheetPart worksheetPart = GetWorksheetPart(spreadsheetDocument, sheetIndex);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;

                foreach (Row row in sheetData.Elements<Row>())
                {
                    List<string> rowData = row.Elements<Cell>().Select(cell => ExcelUtilities.GetCellValue(cell, spreadsheetDocument.WorkbookPart!)).ToList();
                    excelData.Add(rowData);
                }
            }

            return excelData;
        }

        public static List<T> ReadExcelFileByIndex<T>(string filePath, Func<int, List<string>, T?> convertFunctionPointer, int sheetIndex = 0)
        {
            List<T> excelData = new();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorksheetPart worksheetPart = GetWorksheetPart(spreadsheetDocument, sheetIndex);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;

                int rowIndex = 0;
                foreach (Row row in sheetData.Elements<Row>())
                {
                    List<string> rowData = row.Elements<Cell>().Select(cell => ExcelUtilities.GetCellValue(cell, spreadsheetDocument.WorkbookPart!)).ToList();
                    T? convertedData = convertFunctionPointer(rowIndex++, rowData);
                    if (convertedData != null)
                        excelData.Add(convertedData);
                }
            }

            return excelData;
        }

        public static List<Dictionary<string, string>> ReadExcelFileByHeader(string filePath, int sheetIndex = 0)
        {
            List<Dictionary<string, string>> excelData = new();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorksheetPart worksheetPart = GetWorksheetPart(spreadsheetDocument, sheetIndex);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;

                Row headerRow = sheetData.Elements<Row>().FirstOrDefault()!;
                List<string> headers = headerRow.Elements<Cell>().Select(cell => ExcelUtilities.GetCellValue(cell, spreadsheetDocument.WorkbookPart!).ToKey()).ToList();

                foreach (Row row in sheetData.Elements<Row>().Skip(1))
                {
                    Dictionary<string, string> rowData = new();
                    int cellIndex = 0;

                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        string cellValue = ExcelUtilities.GetCellValue(cell, spreadsheetDocument.WorkbookPart!);
                        rowData[headers[cellIndex++]] = cellValue;
                    }

                    excelData.Add(rowData);
                }
            }

            return excelData;
        }

        public static List<T> ReadExcelFileByHeader<T>(string filePath, Func<Dictionary<string, string>, T?> convertFunctionPointer, int sheetIndex = 0) where T : new()
        {
            List<T> excelData = new();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorksheetPart worksheetPart = GetWorksheetPart(spreadsheetDocument, sheetIndex);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;

                Row headerRow = sheetData.Elements<Row>().FirstOrDefault()!;
                List<string> headers = headerRow.Elements<Cell>().Select(cell => ExcelUtilities.GetCellValue(cell, spreadsheetDocument.WorkbookPart!).ToKey()).ToList();

                foreach (Row row in sheetData.Elements<Row>().Skip(1))
                {
                    Dictionary<string, string> rowData = new();
                    int cellIndex = 0;

                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        string cellValue = ExcelUtilities.GetCellValue(cell, spreadsheetDocument.WorkbookPart!);
                        rowData[headers[cellIndex++]] = cellValue;
                    }

                    T? convertedData = convertFunctionPointer(rowData);
                    if (convertedData != null)
                        excelData.Add(convertedData);
                }
            }

            return excelData;
        }

        public static List<T> ReadExcelFileByHeader<T>(string filePath, int sheetIndex = 0, string dateFormat = "dd/MM/yyyy") where T : new()
        {
            return ReadExcelFileByHeader(filePath, rowData => ExcelUtilities.ConvertToObject<T>(rowData, dateFormat), sheetIndex);
        }

        private static WorksheetPart GetWorksheetPart(SpreadsheetDocument document, int sheetIndex)
        {
            WorkbookPart workbookPart = document.WorkbookPart!;
            Sheet sheet = workbookPart.Workbook.GetFirstChild<Sheets>()!.Elements<Sheet>().ElementAt(sheetIndex);
            return (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        }
    }
}
