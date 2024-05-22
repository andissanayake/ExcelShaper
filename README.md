# ExcelShaper

ExcelShaper is a .NET library designed to facilitate reading and shaping Excel files. It provides convenient methods for extracting data from Excel files based on sheet index or header names and offers custom conversion capabilities to convert Excel data into custom types.

## Table of Contents

- [Installation](#installation)
- [Usage](#usage)
  - [Reading Excel Files](#reading-excel-files)
    - [By Index](#by-sheet-index)
    - [By Header Names](#by-header-names)
  - [Custom Conversion](#custom-conversion)
  - [Handling Date Formats](#handling-date-formats)
- [Contributing](#contributing)
- [License](#license)

## Installation

ExcelShaper is available as a NuGet package. You can install it via NuGet Package Manager or .NET CLI.

```bash
dotnet add package ExcelShaper
```

## Usage

### Reading Excel Files

#### By Sheet Index

```csharp
string filePath = "path/to/your/excel/file.xlsx";
var data = Engine.ReadExcelFileByIndex(filePath);
```

#### By Header Names

```csharp
string filePath = "path/to/your/excel/file.xlsx";
var data = Engine.ReadExcelFileByHeader(filePath);
```

### Custom Conversion
```csharp
public class Person
{
    public int Index { get; set; }
    public string FirstName { get; set; } = "";

    //more properties
}

string filePath = "path/to/your/excel/file.xlsx";
var data = Engine.ReadExcelFileByHeader(filePath, (rowData) =>
            {
                return new Person
                {
                    Age = int.Parse(rowData["age"]),
                    Country = rowData["country"],
                    
                    //more properties
                };
            });
```

### Handling Date Formats
```csharp
string filePath = "path/to/your/excel/file.xlsx";
var data = Engine.ReadExcelFileByHeader(filePath, (rowData) =>
{
    // Define your conversion logic here
},sheetIndex : 1,dateFormat : "dd/MM/yyyy");
```
## Contributing

Contributions are welcome! If you encounter any bugs or have suggestions for improvements, feel free to open an issue or submit a pull request.

To contribute to ExcelShaper, follow these steps:

1. Fork the repository.
2. Create a new branch (`git checkout -b feature/improvement`).
3. Make your changes.
4. Commit your changes (`git commit -am 'Add new feature'`).
5. Push to the branch (`git push origin feature/improvement`).
6. Create a new Pull Request.

Please make sure to follow the code style and conventions used in the project and ensure that your changes pass all tests.

## License

This project is licensed under the [MIT License](LICENSE).
