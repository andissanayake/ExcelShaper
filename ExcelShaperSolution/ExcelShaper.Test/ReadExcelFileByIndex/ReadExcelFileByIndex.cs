using ExcelShaperLib;
using FluentAssertions;
using System.Globalization;
namespace ExcelShaperTest.ReadExcelFileByIndex
{
    public class ReadExcelFileByIndexTest
    {
        [Fact]
        public void ReadTest()
        {
            string rootPath = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory))))!;
            string filePath = Path.Combine(rootPath, "ReadExcelFileByIndex", "ReadExcelFileByIndex.xlsx");
            var ret = ExcelShaper.ReadExcelFileByIndex(filePath);
            ret.Should().NotBeNullOrEmpty();
            ret.Count.Should().Be(101);
            ret.ForEach(x => { x.Should().NotBeNullOrEmpty(); x.Count.Should().Be(8); });
        }

        [Fact]
        public void MultipleSheetsTest()
        {
            string rootPath = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory))))!;
            string filePath = Path.Combine(rootPath, "ReadExcelFileByIndex", "ReadExcelFileByIndex.xlsx");
            var ret1 = ExcelShaper.ReadExcelFileByIndex(filePath, 0);
            var ret2 = ExcelShaper.ReadExcelFileByIndex(filePath, 1);

            ret1.Should().NotBeNullOrEmpty();
            ret1.Count.Should().Be(101);
            ret1.ForEach(x => { x.Should().NotBeNullOrEmpty(); x.Count.Should().Be(8); });


            ret2.Should().NotBeNullOrEmpty();
            ret2.Count.Should().Be(21);
            ret2.ForEach(x => { x.Should().NotBeNull(); x.Count.Should().Be(4); });
        }

        [Fact]
        public void ConvertFunctionTest()
        {
            string rootPath = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory))))!;
            string filePath = Path.Combine(rootPath, "ReadExcelFileByIndex", "ReadExcelFileByIndex.xlsx");
            var ret = ExcelShaper.ReadExcelFileByIndex(filePath, (i, rowData) =>
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
            ret.Should().NotBeNullOrEmpty();
            ret.Count.Should().Be(100);
            ret.ForEach(x => { x.Should().NotBeNull(); });

        }

        [Fact]
        public void ReadDateFormatTest()
        {
            string rootPath = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory))))!;
            string filePath = Path.Combine(rootPath, "ReadExcelFileByIndex", "ReadExcelFileByIndex.xlsx");
            var ret = ExcelShaper.ReadExcelFileByIndex(filePath, (i, rowData) =>
            {
                //to avoid first header raw
                if (i > 0)
                {
                    return new Employee
                    {
                        EEID = rowData[0],
                        FullName = rowData[1],
                        JobTitle = rowData[2],
                        Department = rowData[3],
                        BusinessUnit = rowData[4],
                        Gender = rowData[5],
                        Ethnicity = rowData[6],
                        Age = int.Parse(rowData[7]),
                        HireDate = DateTime.FromOADate(int.Parse(rowData[8])),
                        AnnualSalary = int.Parse(rowData[9]),
                        BonusPercentage = decimal.Parse(rowData[10], System.Globalization.NumberStyles.Float),
                        Country = rowData[11],
                        City = rowData[12],
                        ExitDate = string.IsNullOrEmpty(rowData[13]) ? null : DateTime.FromOADate(int.Parse(rowData[13]))
                    };
                }
                return null;
            }, 2);
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