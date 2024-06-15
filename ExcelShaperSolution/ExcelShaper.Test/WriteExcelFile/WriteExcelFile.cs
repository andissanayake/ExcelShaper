using System.Reflection;

namespace ExcelShaper.Test.WriteExcelFile
{
    public class WriteExcelFile
    {
        [Fact]
        public void WriteTest()
        {
            List<Person> people = new List<Person>
            {
                new Person { Name = "John Doe", BirthDate = new DateTime(1990, 5, 10), Address = new Address { Street = "123 Main St", City = "Anytown" } },
                new Person { Name = "Jane Smith", BirthDate = new DateTime(1985, 10, 15), Address = new Address { Street = "456 Elm St", City = "Otherville" } }
            };

            Dictionary<string, string> customHeaders = new Dictionary<string, string>
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

            string filePath = "people.xlsx";
            Engine2.WriteToExcel(filePath, people, customHeaders, complexPropertyConverter);
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
