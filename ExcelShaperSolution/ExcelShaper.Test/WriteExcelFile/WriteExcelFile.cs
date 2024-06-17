using ExcelShaperLib;
using System.Reflection;

namespace ExcelShaperTest.WriteExcelFile
{
    public class WriteExcelFile
    {
        [Fact]
        public void WriteTest()
        {
            string rootPath = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory))))!;
            string filePath = Path.Combine(rootPath, "WriteExcelFile", "WriteExcelFile.xlsx");

            List<Person> people = new()
            {
                new() { Name = "John Doe", BirthDate = new DateTime(1990, 5, 10), Address = new Address { Street = "123 Main St", City = "Anytown" } },
                new() { Name = "Jane Smith", BirthDate = new DateTime(1985, 10, 15), Address = new Address { Street = "456 Elm St", City = "Otherville" } }
            };

            Dictionary<string, string> customHeaders = new()
            {
                { "Name", "Full Name" },
                { "BirthDate", "Date of Birth" },
                { "Address", "Home Address" }
            };

            Func<Person, PropertyInfo, string> complexPropertyConverter = (person, property) =>
            {
                if (property.Name == "BirthDate")
                {
                    return ((DateTime)property.GetValue(person)).ToString("yyyy-MM-dd");
                }
                else if (property.Name == "Address")
                {
                    Address address = (Address)property.GetValue(person);
                    return $"{address.Street}, {address.City}";
                }
                else
                {
                    return property.GetValue(person)?.ToString() ?? string.Empty;
                }
            };

            ExcelShaper.WriteToExcel(filePath, people, customHeaders, complexPropertyConverter);
        }

    }
    public class Person
    {
        public string Name { get; set; }
        public DateTime BirthDate { get; set; }
        public Address Address { get; set; }
    }

    public class Address
    {
        public string Street { get; set; }
        public string City { get; set; }
    }
}
