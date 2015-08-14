using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelLibrary;

namespace ExcelPOC
{
    class Program
    {
        static void Main(string[] args)
        {
            //Method();
            try
            {
                List<Package> packages =
                    new List<Package>
                        { new Package { Company = "Coho Vineyard", Weight = 25.2,
                              TrackingNumber = 89453312L,
                              DateOrder = DateTime.Today, HasCompleted = false },
                          new Package { Company = "Lucerne Publishing", Weight = 18.7,
                              TrackingNumber = 89112755L,
                              DateOrder = DateTime.Today, HasCompleted = false },
                          new Package { Company = "Wingtip Toys", Weight = 6.0,
                              TrackingNumber = 299456122L,
                              DateOrder = DateTime.Today, HasCompleted = false },
                          new Package { Company = "Adventure Works", Weight = 33.8,
                              TrackingNumber = 4665518773L,
                              DateOrder =  DateTime.Today.AddDays(-4),
                              HasCompleted = true },
                          new Package { Company = "Test Works", Weight = 35.8,
                              TrackingNumber = 4665518774L,
                              DateOrder =  DateTime.Today.AddDays(-2),
                              HasCompleted = true },
                          new Package { Company = "Good Works", Weight = 48.8,
                              TrackingNumber = 4665518775L,
                              DateOrder =  DateTime.Today.AddDays(-1), HasCompleted = true },

                        };

                var tuples = new[]
                {
                    new Tuple<string, bool, int>("Le Lorem Ipsum est simplement du faux texte employé dans la composition et la mise en page avant impression. Le Lorem Ipsum est le faux texte standard de l'imprimerie depuis les années 1500, quand un peintre anonyme assembla ensemble des morceaux de texte pour réaliser un livre spécimen de polices de texte. Il n'a pas fait que survivre cinq siècles, mais s'est aussi adapté à la bureautique informatique, sans que son contenu n'en soit modifié. Il a été popularisé dans les années 1960 grâce à la vente de feuilles Letraset contenant des passages du Lorem Ipsum, et, plus récemment, par son inclusion dans des applications de mise en page de texte, comme Aldus PageMaker.", false, 5),
                    new Tuple<string, bool, int>("beta", true, 11),
                    new Tuple<string, bool, int>("epsilon", false, 10),
                    new Tuple<string, bool, int>("gamma", true, 1)
                };

                var excel = new Excel();

                excel.AddSheet<Package>("SheetName")
                    .Map("A", "Alphabet", p => p.DateOrder, "yyyy-MM-dd")
                    .Map("B", "Bim", p => p.Company)
                    .Map("C", p => p.Weight)
                    .Map("D", "Damned", p => p.TrackingNumber)
                    .Map("E", "Epsilon", p => (int)p.Weight % 2 == 0)
                    .SetData(packages);

                excel.AddSheet<Tuple<string, bool, int>>("TupleSheet")
                    .Map("A", t => t.Item1)
                    .Map("B", t => t.Item3)
                    .Map("C", t => t.Item2)
                    .Map("D", t => t.Item3 * (t.Item2 ? 2 : 3))
                    .SetData(tuples);


                excel.SaveTo("grab.xlsx");
                

                Console.WriteLine("Completed");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            Console.Read();
        }

        public static void Method()
        {
            Console.WriteLine("Creating document");
            using (var spreadsheet = SpreadsheetDocument.Create("output.xlsx", SpreadsheetDocumentType.Workbook))
            {
                Console.WriteLine("Creating workbook");
                spreadsheet.AddWorkbookPart();
                spreadsheet.WorkbookPart.Workbook = new Workbook();
                Console.WriteLine("Creating worksheet");
                var wsPart = spreadsheet.WorkbookPart.AddNewPart<WorksheetPart>();
                wsPart.Worksheet = new Worksheet();

                var stylesPart = spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet();

                Console.WriteLine("Creating styles");

                // blank font list
                stylesPart.Stylesheet.Fonts = new Fonts();
                stylesPart.Stylesheet.Fonts.Count = 1;
                stylesPart.Stylesheet.Fonts.AppendChild(new Font());

                // create fills
                stylesPart.Stylesheet.Fills = new Fills();

                // create a solid red fill
                var solidRed = new PatternFill() { PatternType = PatternValues.Solid };
                solidRed.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FFFF0000") }; // red fill
                solidRed.BackgroundColor = new BackgroundColor { Indexed = 64 };

                stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
                stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
                stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = solidRed });
                stylesPart.Stylesheet.Fills.Count = 3;

                // blank border list
                stylesPart.Stylesheet.Borders = new Borders();
                stylesPart.Stylesheet.Borders.Count = 1;
                stylesPart.Stylesheet.Borders.AppendChild(new Border());

                // blank cell format list
                stylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats();
                stylesPart.Stylesheet.CellStyleFormats.Count = 1;
                stylesPart.Stylesheet.CellStyleFormats.AppendChild(new CellFormat());

                // cell format list
                stylesPart.Stylesheet.CellFormats = new CellFormats();
                // empty one for index 0, seems to be required
                stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat());
                // cell format references style format 0, font 0, border 0, fill 2 and applies the fill
                stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 0, BorderId = 0, FillId = 2, ApplyFill = true }).AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Center });
                stylesPart.Stylesheet.CellFormats.Count = 2;

                stylesPart.Stylesheet.Save();

                Console.WriteLine("Creating sheet data");
                var sheetData = wsPart.Worksheet.AppendChild(new SheetData());

                Console.WriteLine("Adding rows / cells...");

                var row = sheetData.AppendChild(new Row());
                row.AppendChild(new Cell() { CellValue = new CellValue("This"), DataType = CellValues.String });
                row.AppendChild(new Cell() { CellValue = new CellValue("is"), DataType = CellValues.String });
                row.AppendChild(new Cell() { CellValue = new CellValue("a"), DataType = CellValues.String });
                row.AppendChild(new Cell() { CellValue = new CellValue("test."), DataType = CellValues.String });

                sheetData.AppendChild(new Row());

                row = sheetData.AppendChild(new Row());
                row.AppendChild(new Cell() { CellValue = new CellValue("Value:"), DataType = CellValues.String });
                row.AppendChild(new Cell() { CellValue = new CellValue("123"), DataType = CellValues.Number });
                row.AppendChild(new Cell() { CellValue = new CellValue("Formula:"), DataType = CellValues.String });
                // style index = 1, i.e. point at our fill format
                row.AppendChild(new Cell() { CellFormula = new CellFormula("B3"), DataType = CellValues.Number, StyleIndex = 1 });

                Console.WriteLine("Saving worksheet");
                wsPart.Worksheet.Save();

                Console.WriteLine("Creating sheet list");
                var sheets = spreadsheet.WorkbookPart.Workbook.AppendChild(new Sheets());
                sheets.AppendChild(new Sheet() { Id = spreadsheet.WorkbookPart.GetIdOfPart(wsPart), SheetId = 1, Name = "Test" });

                Console.WriteLine("Saving workbook");
                spreadsheet.WorkbookPart.Workbook.Save();

                Console.WriteLine("Done.");
            }
        }

    }



    public class Package
    {
        public string Company { get; set; }
        public double Weight { get; set; }
        public long TrackingNumber { get; set; }
        public DateTime DateOrder { get; set; }
        public bool HasCompleted { get; set; }
    }
}
