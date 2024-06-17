using ExcelShaperLib;
using FluentAssertions;
using System.Globalization;
namespace ExcelShaperTest.ReadExcelFileByHeader
{
    public class ReadExcelFileByHeaderTests
    {
        [Fact]
        public void ReadTest()
        {
            string rootPath = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory))))!;
            string filePath = Path.Combine(rootPath, "ReadExcelFileByHeader", "ReadExcelFileByHeader.xlsx");
            var ret = ExcelShaper.ReadExcelFileByHeader(filePath);
            ret.Should().NotBeNullOrEmpty();
            ret.Count.Should().Be(100);
            ret.ForEach(x => { x.Should().NotBeNullOrEmpty(); x.Count.Should().Be(8); });
            ret.ForEach(x => x.Where(s => s.Key == "firstname").Count().Should().Be(1));
        }

        [Fact]
        public void ConvertFunctionTest()
        {
            string rootPath = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory))))!;
            string filePath = Path.Combine(rootPath, "ReadExcelFileByHeader", "ReadExcelFileByHeader.xlsx");
            var ret = ExcelShaper.ReadExcelFileByHeader(filePath, (rowData) =>
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
            ret.Should().NotBeNullOrEmpty();
            ret.Count.Should().Be(100);
            ret.ForEach(x => { x.Should().NotBeNull(); });

        }

        [Fact]
        public void ConvertFunctionInbuildTest()
        {
            string rootPath = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory))))!;
            string filePath = Path.Combine(rootPath, "ReadExcelFileByHeader", "ReadExcelFileByHeader.xlsx");
            var ret = ExcelShaper.ReadExcelFileByHeader<Person>(filePath);
            ret.Should().NotBeNullOrEmpty();
            ret.Count.Should().Be(100);
            ret.ForEach(x => { x.Should().NotBeNull(); });

        }

        [Fact]
        public void ConvertFunctionInbuildWithSheetIndexTest()
        {
            string rootPath = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory))))!;
            string filePath = Path.Combine(rootPath, "ReadExcelFileByHeader", "ReadExcelFileByHeader.xlsx");
            var ret = ExcelShaper.ReadExcelFileByHeader<Employee>(filePath, 2);
            ret.Should().NotBeNullOrEmpty();
            ret.Count.Should().Be(1000);
            ret.ForEach(x => { x.Should().NotBeNull(); });

            var exitedCount = ret.Where(p => p.ExitDate != null).Count();
            exitedCount.Should().Be(85);
            var directorCount = ret.Where(p => p.JobTitle == "Director").Count();
            directorCount.Should().Be(121);
        }
    }
    public class Person
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
    public class Employee
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
