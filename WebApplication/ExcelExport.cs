using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication
{
    public class User
    {
        public int Id { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
    }
    public class ExcelExport
    {
        private List<User> Users = new List<User>
        {
        new User()
        {
            Id = 1, UserName = "Test User 1 Test User 1 Test User 1 Test User 1", Password = "Test Password 1 Test Password 1 Test Password 1 Test Password 1"
        },
        new User()
        {
            Id = 2, UserName = "Test User 2", Password = "Test Password 2"
        },
        new User()
        {
            Id = 3, UserName = "Test User 3", Password = "Test Password 3"
        }
        };

        public byte[] CreateExcelDoc()
        {
            var stream = new MemoryStream();
            SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                // Adding style
                WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylePart.Stylesheet = GenerateStylesheet();
                stylePart.Stylesheet.Save();

                Columns worksheetColumns = CreateWorksheetColumns();
                worksheetPart.Worksheet.AppendChild(worksheetColumns);
                workbookPart.Workbook.Save();

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Users" };

                sheets.Append(sheet);

                workbookPart.Workbook.Save();


                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                // Constructing header
                Row row = new Row();

                row.Append(
                    ConstructCell("Id", CellValues.String, 2),
                    ConstructCell("Name", CellValues.String, 2),
                    ConstructCell("Birth Date", CellValues.String, 2));

                // Insert the header row to the Sheet Data
                sheetData.AppendChild(row);

                // Inserting each employee
                foreach (var user in Users)
                {
                    row = new Row();

                row.Append(
                    ConstructCell(user.Id.ToString(), CellValues.Number, 1),
                    ConstructCell(user.UserName, CellValues.String, 1),
                    ConstructCell(user.Password, CellValues.String, 1));

                    sheetData.AppendChild(row);
                }

                document.Clone(stream);
                worksheetPart.Worksheet.Save();

            return stream.ToArray();
        }

        private Cell ConstructCell(string value, CellValues dataType, uint styleIndex = 0)
        {
            var cell = new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType),
                StyleIndex = styleIndex
            };
            return cell;
        }

        private Columns CreateWorksheetColumns()
        {
            // define a new columns object
            Columns workSheetColumns = new Columns();

            // invoice number column
            Column col = new Column();
            col.Width = DoubleValue.FromDouble(5.0);
            col.Min = UInt32Value.FromUInt32((UInt32)1);
            col.Max = col.Min;
            col.CustomWidth = BooleanValue.FromBoolean(true);
            workSheetColumns.Append(col);

            // date column
            col = new Column();
            col.Width = DoubleValue.FromDouble(25.0);
            col.Min = UInt32Value.FromUInt32((UInt32)2);
            col.Max = col.Min;
            col.CustomWidth = BooleanValue.FromBoolean(true);
            workSheetColumns.Append(col);

            // first name column
            col = new Column();
            col.Width = DoubleValue.FromDouble(25);
            col.Min = UInt32Value.FromUInt32((UInt32)3);
            col.Max = col.Min;
            col.CustomWidth = BooleanValue.FromBoolean(true);
            workSheetColumns.Append(col);

            return workSheetColumns;
        }

        private Stylesheet GenerateStylesheet()
        {
            Stylesheet styleSheet = new Stylesheet();

            Fonts fonts = new Fonts(
                new Font( // Index 0 - default
                    new FontSize() { Val = 10 },
                    new Color() { Rgb = "000000" }

                ),
                new Font( // Index 1 - header
                    new FontSize() { Val = 14 },
                    new Bold(),
                    new Color() { Rgb = "ffffff" }

                ));

            Fills fills = new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
                    new Fill(new PatternFill(new ForegroundColor { Rgb = "9fd6fb" }) { PatternType = PatternValues.Solid }), 
                    new Fill(new PatternFill(new ForegroundColor { Rgb = "06f929" }) { PatternType = PatternValues.Solid })

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
                    new CellFormat(), //Index 0 - default
                    new CellFormat() { FontId = 0, FillId = 2, BorderId = 1, Alignment = new Alignment { WrapText = true } }, // body
                    new CellFormat() { FontId = 1, FillId = 3, BorderId = 1, Alignment = new Alignment { WrapText = true } } // header
                );

            styleSheet = new Stylesheet(fonts, fills, borders, cellFormats);

            return styleSheet;
        }
    }
}
