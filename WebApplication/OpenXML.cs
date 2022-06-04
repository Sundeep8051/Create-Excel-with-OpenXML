//Source Code http://www.dispatchertimer.com/tutorial/how-to-create-an-excel-file-in-net-using-openxml-part-3-add-stylesheet-to-the-spreadsheet/

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLSample
{
    public class Report
    {
        public void CreateExcelDoc(string fileName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                // Adding style
                WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylePart.Stylesheet = GenerateStylesheet();
                stylePart.Stylesheet.Save();

                // Setting up columns
                Columns columns = new Columns(
                        new Column // Id column
                        {
                            Min = 1,
                            Max = 1,
                            Width = 4,
                            CustomWidth = true
                        },
                        new Column // Name and Birthday columns
                        {
                            Min = 2,
                            Max = 3,
                            Width = 15,
                            CustomWidth = true
                        },
                        new Column // Salary column
                        {
                            Min = 4,
                            Max = 4,
                            Width = 8,
                            CustomWidth = true
                        });

                worksheetPart.Worksheet.AppendChild(columns);

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Employees" };

                sheets.Append(sheet);

                workbookPart.Workbook.Save();

                List<Employee> employees = Employees.EmployeesList;

                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                // Constructing header
                Row row = new Row();

                row.Append(
                    ConstructCell("Id", CellValues.String, 2),
                    ConstructCell("Name", CellValues.String, 2),
                    ConstructCell("Birth Date", CellValues.String, 2),
                    ConstructCell("Salary", CellValues.String, 2));

                // Insert the header row to the Sheet Data
                sheetData.AppendChild(row);

                // Inserting each employee
                foreach (var employee in employees)
                {
                    row = new Row();

                    row.Append(
                        ConstructCell(employee.Id.ToString(), CellValues.Number, 1),
                        ConstructCell(employee.Name, CellValues.String, 1),
                        ConstructCell(employee.DOB.ToString("yyyy/MM/dd"), CellValues.String, 1),
                        ConstructCell(employee.Salary.ToString(), CellValues.Number, 1));

                    sheetData.AppendChild(row);
                }

                worksheetPart.Worksheet.Save();
            }
        }

        private Cell ConstructCell(string value, CellValues dataType, uint styleIndex = 0)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType),
                StyleIndex = styleIndex
            };
        }

        private Stylesheet GenerateStylesheet()
        {
            Stylesheet styleSheet = null;

            Fonts fonts = new Fonts(
                new Font( // Index 0 - default
                    new FontSize() { Val = 10 }

                ),
                new Font( // Index 1 - header
                    new FontSize() { Val = 10 },
                    new Bold(),
                    new Color() { Rgb = "FFFFFF" }

                ));

            Fills fills = new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "66666666" } })
                    { PatternType = PatternValues.Solid }) // Index 2 - header
                );

            Borders borders = new Borders(
                    new Border(), // index 0 default
                    new Border( // index 1 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                );

            CellFormats cellFormats = new CellFormats(
                    new CellFormat(), // default
                    new CellFormat { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }, // body
                    new CellFormat { FontId = 1, FillId = 2, BorderId = 1, ApplyFill = true } // header
                );

            styleSheet = new Stylesheet(fonts, fills, borders, cellFormats);

            return styleSheet;
        }
    }

    public class Employee
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public DateTime DOB { get; set; }
        public decimal Salary { get; set; }
    }

    public sealed class Employees
    {
        static List<Employee> _employees;
        const int COUNT = 15;

        public static List<Employee> EmployeesList
        {
            private set { }
            get
            {
                return _employees;
            }
        }

        static Employees()
        {
            Initialize();
        }

        private static void Initialize()
        {
            _employees = new List<Employee>();

            Random random = new Random();

            for (int i = 0; i < COUNT; i++)
            {
                _employees.Add(new Employee()
                {
                    Id = i,
                    Name = "Employee " + i,
                    DOB = new DateTime(1999, 1, 1).AddMonths(i),
                    Salary = random.Next(100, 10000)
                });
            }
        }
    }
}