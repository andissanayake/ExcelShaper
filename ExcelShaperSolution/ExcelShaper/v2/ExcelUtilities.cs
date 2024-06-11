using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Reflection;

namespace ExcelShaper
{
    public static class ExcelUtilities
    {
        public static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            string cellValue = cell.CellValue!.Text;
            if (cell.DataType?.Value == CellValues.SharedString)
            {
                int sharedStringId = int.Parse(cell.CellValue.Text);
                SharedStringTablePart sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()!;
                cellValue = sharedStringTablePart.SharedStringTable.ElementAt(sharedStringId).InnerText;
            }
            return cellValue;
        }

        public static string ToKey(this string str)
        {
            return str.Trim().ToLower().Replace(" ", "");
        }

        public static T ConvertToObject<T>(Dictionary<string, string> rowData, string dateFormat = "dd/MM/yyyy") where T : new()
        {
            if (rowData == null)
                throw new ArgumentNullException(nameof(rowData));

            T obj = new();
            var propertyMap = MapProperties<T>(rowData.Keys.ToList());

            foreach (var pair in rowData)
            {
                if (propertyMap.TryGetValue(pair.Key, out PropertyInfo? property))
                {
                    object value = property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(DateTime?)
                        ? ConvertToDate(pair.Value, dateFormat)
                        : Convert.ChangeType(pair.Value, property.PropertyType);

                    property.SetValue(obj, value);
                }
            }

            return obj;
        }

        private static object ConvertToDate(string value, string dateFormat)
        {
            if (string.IsNullOrEmpty(value))
                return default(DateTime?);

            if (int.TryParse(value, out int intValue))
                return DateTime.FromOADate(intValue);

            if (DateTime.TryParseExact(value, dateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateTimeValue))
                return dateTimeValue;

            throw new InvalidCastException($"Unable to convert '{value}' to DateTime");
        }

        private static Dictionary<string, PropertyInfo> MapProperties<T>(List<string> columnNames)
        {
            var properties = typeof(T).GetProperties();
            var propertyMap = new Dictionary<string, PropertyInfo>();

            foreach (string columnName in columnNames)
            {
                var property = properties.FirstOrDefault(p => p.Name.ToKey().Equals(columnName, StringComparison.OrdinalIgnoreCase));
                if (property != null)
                {
                    propertyMap.Add(columnName, property);
                }
            }

            return propertyMap;
        }
    }
}
