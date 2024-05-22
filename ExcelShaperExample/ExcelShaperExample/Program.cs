using ExcelShaper;
using System.Globalization;
using System.Text.Json;

namespace ExcelShaperExample
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string rootPath = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory))))!;
            string filePath = Path.Combine(rootPath, "ex1.xlsx");

            var ret1 = Engine.ReadExcelFileByIndex(filePath);
            Console.WriteLine("1. Read with index sheet 1 -> ");
            Console.WriteLine(JsonSerializer.Serialize(ret1.Take(3)));
            Console.WriteLine();

            var ret2 = Engine.ReadExcelFileByIndex(filePath, 1);
            Console.WriteLine("2. Read with index sheet 2 -> ");
            Console.WriteLine(JsonSerializer.Serialize(ret2.Take(3)));
            Console.WriteLine();

            var ret3 = Engine.ReadExcelFileByIndex(filePath, 2);
            Console.WriteLine("3. Read with index sheet 3 -> ");
            Console.WriteLine(JsonSerializer.Serialize(ret3.Take(3)));
            Console.WriteLine();

            var ret4 = Engine.ReadExcelFileByHeader(filePath);
            Console.WriteLine("4. Read with header sheet 1 -> ");
            Console.WriteLine(JsonSerializer.Serialize(ret4.Take(3)));
            Console.WriteLine();

            var ret5 = Engine.ReadExcelFileByHeader(filePath, 1);
            Console.WriteLine("5. Read with header sheet 2 -> ");
            Console.WriteLine(JsonSerializer.Serialize(ret5.Take(3)));
            Console.WriteLine();

            var ret6 = Engine.ReadExcelFileByHeader(filePath, 2);
            Console.WriteLine("6. Read with header sheet 3 -> ");
            Console.WriteLine(JsonSerializer.Serialize(ret6.Take(3)));
            Console.WriteLine();

            var ret7 = Engine.ReadExcelFileByIndex(filePath, (i, rowData) =>
            {
                //to avoid first header raw
                if (i > 0)
                {
                    return new Person
                    {
                        Age = int.Parse(rowData[5]),
                        Country = rowData[4],
                        Date = DateTime.ParseExact(rowData[6], "dd/MM/yyyy", CultureInfo.InvariantCulture),
                        FirstName = rowData[1],
                        Gender = rowData[3],
                        Id = int.Parse(rowData[7]),
                        LastName = rowData[2],
                        Index = int.Parse(rowData[0]),
                    };
                }
                return null;
            });
            Console.WriteLine("7. Read with index sheet 1 with mapping function -> ");
            Console.WriteLine(JsonSerializer.Serialize(ret7.Take(3)));
            Console.WriteLine();

            var ret8 = Engine.ReadExcelFileByHeader(filePath, (rowData) =>
            {
                return new Person
                {
                    Age = int.Parse(rowData["age"]),
                    Country = rowData["country"],
                    Date = DateTime.ParseExact(rowData["date"], "dd/MM/yyyy", CultureInfo.InvariantCulture),
                    FirstName = rowData["firstname"],
                    Gender = rowData["gender"],
                    Id = int.Parse(rowData["id"]),
                    LastName = rowData["lastname"],
                    Index = int.Parse(rowData["index"]),
                };
            });
            Console.WriteLine("8. Read with header sheet 1 with mapping function -> ");
            Console.WriteLine(JsonSerializer.Serialize(ret8.Take(3)));
            Console.WriteLine();

            string filePath1 = Path.Combine(rootPath, "ex2.xlsx");
            var ret9 = Engine.ReadExcelFileByHeader<Person>(filePath);
            Console.WriteLine("9. Read with header sheet 1 with inbuild mapping function -> ");
            Console.WriteLine(JsonSerializer.Serialize(ret9.Take(3)));
            Console.WriteLine();

            Console.ReadLine();
        }
    }
    internal class Person
    {
        public int Index { get; set; }
        public string FirstName { get; set; } = "";
        public string LastName { get; set; } = "";
        public string Gender { get; set; } = "";
        public string Country { get; set; } = "";
        public int Age { get; set; }
        public DateTime Date { get; set; }
        public int Id { get; set; }
    }
    internal class Employee
    {
        public string EEID { get; set; } = "";
        public string FullName { get; set; } = "";
        public string JobTitle { get; set; } = "";
        public string Department { get; set; } = "";
        public string BusinessUnit { get; set; } = "";
        public string Gender { get; set; } = "";
        public string Ethnicity { get; set; } = "";
        public int Age { get; set; }
        public DateTime HireDate { get; set; }
        public int AnnualSalary { get; set; }
        public decimal BonusPercentage { get; set; }
        public string Country { get; set; } = "";
        public string City { get; set; } = "";
        public DateTime? ExitDate { get; set; }
    }
}
