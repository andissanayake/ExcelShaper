using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Reflection;

namespace ExcelShaperLib
{
    public static class Engine
    {
        /// <summary>
        /// Reads an Excel file and returns the data as a list of lists of strings, where each inner list represents a row of data.
        /// </summary>
        /// <param name="filePath">The path to the Excel file.</param>
        /// <param name="sheetIndex">The index of the sheet to read (default is 0).</param>
        /// <returns>A list of lists of strings representing the data in the Excel file.</returns>
        public static List<List<string>> ReadExcelFileByIndex(string filePath, int sheetIndex = 0)
        {
            List<List<string>> excelData = new();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart!;
                IEnumerable<Sheet> sheets = spreadsheetDocument.WorkbookPart!.Workbook!.GetFirstChild<Sheets>()!.Elements<Sheet>();
                string relationshipId = sheets.ElementAt(sheetIndex).Id!.Value!;
                WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet worksheet = worksheetPart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>()!;

                for (int i = 0; i < sheetData.Elements<Row>().Count(); i++)
                {
                    Row row = sheetData.Elements<Row>().ElementAt(i);
                    List<string> rowData = new();

                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        string cellValue = GetCellValue(cell, workbookPart);
                        rowData.Add(cellValue);
                    }

                    excelData.Add(rowData);
                }
            }

            return excelData;
        }

        /// <summary>
        /// Reads an Excel file and converts each row of data using a custom conversion function.
        /// </summary>
        /// <typeparam name="T">The type to convert each row of data to.</typeparam>
        /// <param name="filePath">The path to the Excel file.</param>
        /// <param name="convertFunctionPointer">A function pointer to a function that converts a row of data to type T.
        /// The function signature should be: T? ConvertFunction(int rowIndex, List&lt;string&gt; rowData), 
        /// where rowIndex is the index of the current row being converted, and rowData is a list of cell values for that row.</param>
        /// <param name="sheetIndex">The index of the sheet to read (default is 0).</param>
        /// <returns>A list of objects of type T representing the converted data.</returns>
        public static List<T> ReadExcelFileByIndex<T>(string filePath, Func<int, List<string>, T?> convertFunctionPointer, int sheetIndex = 0)
        {
            List<T> excelData = new();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart!;
                IEnumerable<Sheet> sheets = spreadsheetDocument.WorkbookPart!.Workbook!.GetFirstChild<Sheets>()!.Elements<Sheet>();
                string relationshipId = sheets.ElementAt(sheetIndex).Id!.Value!;
                WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet worksheet = worksheetPart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>()!;

                for (int i = 0; i < sheetData.Elements<Row>().Count(); i++)
                {
                    Row row = sheetData.Elements<Row>().ElementAt(i);
                    List<string> rowData = new();

                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        string cellValue = GetCellValue(cell, workbookPart);
                        rowData.Add(cellValue);
                    }

                    T? convertedData = convertFunctionPointer(i, rowData);
                    if (convertedData != null)
                        excelData.Add(convertedData);
                }
            }

            return excelData;
        }

        /// <summary>
        /// Reads an Excel file and returns the data as a list of dictionaries, where each dictionary represents a row of data
        /// with column headers as keys and cell values as values.
        /// </summary>
        /// <param name="filePath">The path to the Excel file.</param>
        /// <param name="sheetIndex">The index of the sheet to read (default is 0).</param>
        /// <returns>A list of dictionaries representing the data in the Excel file.</returns>
        public static List<Dictionary<string, string>> ReadExcelFileByHeader(string filePath, int sheetIndex = 0)
        {
            List<Dictionary<string, string>> excelData = new();
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart!;
                IEnumerable<Sheet> sheets = spreadsheetDocument.WorkbookPart!.Workbook!.GetFirstChild<Sheets>()!.Elements<Sheet>();
                string relationshipId = sheets.ElementAt(sheetIndex).Id!.Value!;
                WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet worksheet = worksheetPart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>()!;

                for (int i = 1; i < sheetData.Elements<Row>().Count(); i++)
                {
                    Row row = sheetData.Elements<Row>().ElementAt(i);
                    Dictionary<string, string> rowData = new();

                    for (int j = 0; j < row.Elements<Cell>().Count(); j++)
                    {
                        string keyCellValue = GetCellValue(sheetData.Elements<Row>().ElementAt(0).Elements<Cell>().ElementAt(j), workbookPart).Trim().Replace(" ", "").ToLower();
                        string cellValue = GetCellValue(row.Elements<Cell>().ElementAt(j), workbookPart);
                        rowData.Add(keyCellValue, cellValue);
                    }
                    excelData.Add(rowData);
                }
            }

            return excelData;
        }

        /// <summary>
        /// Reads an Excel file and converts each row of data to type T using a custom conversion function.
        /// </summary>
        /// <typeparam name="T">The type to convert each row of data to.</typeparam>
        /// <param name="filePath">The path to the Excel file.</param>
        /// <param name="convertFunctionPointer">A function pointer to a function that converts a row of data to type T.
        /// The function should have the following signature: T? ConvertFunction(Dictionary&lt;string, string&gt; rowData),
        /// where rowData is a dictionary representing a row of data with column headers as keys and cell values as values.</param>
        /// <param name="sheetIndex">The index of the sheet to read (default is 0).</param>
        /// <returns>A list of objects of type T representing the converted data.</returns>
        public static List<T> ReadExcelFileByHeader<T>(string filePath, Func<Dictionary<string, string>, T?> convertFunctionPointer, int sheetIndex = 0) where T : new()
        {
            List<T> excelData = new();
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart!;
                IEnumerable<Sheet> sheets = spreadsheetDocument.WorkbookPart!.Workbook!.GetFirstChild<Sheets>()!.Elements<Sheet>();
                string relationshipId = sheets.ElementAt(sheetIndex).Id!.Value!;
                WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet worksheet = worksheetPart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>()!;

                for (int i = 1; i < sheetData.Elements<Row>().Count(); i++)
                {
                    Row row = sheetData.Elements<Row>().ElementAt(i);
                    Dictionary<string, string> rowData = new();

                    for (int j = 0; j < row.Elements<Cell>().Count(); j++)
                    {
                        string keyCellValue = GetCellValue(sheetData.Elements<Row>().ElementAt(0).Elements<Cell>().ElementAt(j), workbookPart).ToKey();
                        string cellValue = GetCellValue(row.Elements<Cell>().ElementAt(j), workbookPart);
                        rowData.Add(keyCellValue, cellValue);
                    }
                    T? convertedData = convertFunctionPointer(rowData);

                    if (convertedData != null)
                        excelData.Add(convertedData);
                }
            }

            return excelData;
        }

        /// <summary>
        /// Reads an Excel file and converts each row of data to type T using a default conversion function.
        /// </summary>
        /// <typeparam name="T">The type to convert each row of data to.</typeparam>
        /// <param name="filePath">The path to the Excel file.</param>
        /// <param name="sheetIndex">The index of the sheet to read (default is 0).</param>
        /// <param name="dateFormat">If date column exists in file format (default is "dd/MM/yyyy").</param>
        /// <returns>A list of objects of type T representing the converted data.</returns>
        public static List<T> ReadExcelFileByHeader<T>(string filePath, int sheetIndex = 0, string dateFormat = "dd/MM/yyyy") where T : new()
        {
            List<T> excelData = new();
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart!;
                IEnumerable<Sheet> sheets = spreadsheetDocument.WorkbookPart!.Workbook!.GetFirstChild<Sheets>()!.Elements<Sheet>();
                string relationshipId = sheets.ElementAt(sheetIndex).Id!.Value!;
                WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet worksheet = worksheetPart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>()!;

                for (int i = 1; i < sheetData.Elements<Row>().Count(); i++)
                {
                    Row row = sheetData.Elements<Row>().ElementAt(i);
                    Dictionary<string, string> rowData = new();

                    for (int j = 0; j < row.Elements<Cell>().Count(); j++)
                    {
                        string keyCellValue = GetCellValue(sheetData.Elements<Row>().ElementAt(0).Elements<Cell>().ElementAt(j), workbookPart).ToKey();
                        string cellValue = GetCellValue(row.Elements<Cell>().ElementAt(j), workbookPart);
                        rowData.Add(keyCellValue, cellValue);
                    }
                    T convertedData = ConvertFunction<T>(rowData, dateFormat);
                    if (convertedData != null)
                        excelData.Add(convertedData);
                }
            }

            return excelData;
        }
        private static T ConvertFunction<T>(Dictionary<string, string> rowData, string dateFormat = "dd/MM/yyyy") where T : new()
        {
            if (rowData == null)
            {
                throw new ArgumentNullException(nameof(rowData));
            }

            T obj = new();
            var columnNames = rowData.Keys.ToList();
            var propertyMap = MapProperties<T>(columnNames);

            foreach (var pair in rowData)
            {
                if (propertyMap.TryGetValue(pair.Key, out PropertyInfo? property))
                {
                    object value;
                    if (property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(DateTime?))
                    {
                        if (string.IsNullOrEmpty(pair.Value))
                        {
                            value = default!;
                        }
                        else if (int.TryParse(pair.Value, out int intValue))
                        {
                            value = DateTime.FromOADate(intValue);
                        }
                        else if (DateTime.TryParseExact(pair.Value, dateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateTimeValue))
                        {
                            value = dateTimeValue;
                        }
                        else
                        {
                            throw new InvalidCastException($"Unable to convert '{pair.Value}' to DateTime for property '{property.Name}'");
                        }
                    }
                    else
                    {
                        value = Convert.ChangeType(pair.Value, property.PropertyType);
                    }

                    property.SetValue(obj, value);
                }
            }

            return obj;
        }
        private static Dictionary<string, PropertyInfo> MapProperties<T>(List<string> columnNames)
        {
            var properties = typeof(T).GetProperties();
            var propertyMap = new Dictionary<string, PropertyInfo>();

            for (int i = 0; i < columnNames.Count; i++)
            {
                string columnName = columnNames[i];
                var property = properties.FirstOrDefault(p => p.Name.ToKey().Equals(columnName.ToKey(), StringComparison.OrdinalIgnoreCase));
                if (property != null)
                {
                    propertyMap.Add(property.Name.ToKey(), property);
                }
            }

            return propertyMap;
        }
        private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            string cellValue = cell.CellValue!.Text;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                int sharedStringId = int.Parse(cell.CellValue.Text);
                SharedStringTablePart sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()!;
                if (sharedStringTablePart != null)
                {
                    cellValue = sharedStringTablePart.SharedStringTable.ElementAt(sharedStringId).InnerText;
                }
            }
            return cellValue;
        }
        //private static string ToKey(this string str)
        //{
        //    return str.Trim().ToLower().Replace(" ", "");
        //}
    }
}
