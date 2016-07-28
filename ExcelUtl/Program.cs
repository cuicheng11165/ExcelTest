using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelUtl;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelTest
{
    public class TestEntity
    {
        [SheetColumn("")]
        public string TestString { set; get; }

        [SheetColumn("")]
        public int TestInt32 { set; get; }

        [SheetColumn("")]
        public double TestDouble { set; get; }

        [SheetColumn("")]
        public bool TestBoolean { set; get; }

        [SheetColumn("")]
        public DateTime TestDateTime { set; get; }
    }

    public class RowAtt : Attribute
    {
    }

    class Program
    {
        static void Main(string[] args)
        {
            //var entity = new TestEntity { TestString = "aqadsdasd", TestInt32 = 100, TestDouble = 1.2122, TestBoolean = true, TestDateTime = DateTime.Now };
            var entity = new TestEntity { TestString = "aqadsdasd", TestInt32 = 100, TestDouble = 1.2122, TestBoolean = true, TestDateTime = DateTime.Now };
            var list = new List<TestEntity>() { entity };

            var document = new SpreadsheetImport("d:\\test2.xlsx", "sheet1");

            document.AddElements(list);

            document.Dispose();

            return;



            CreateExcel("d:\\test.xlsx", "name1", new string[][] { new string[] { "Column1", "Column2" }, new string[] { "222", "333" }, new string[] { "111", "2222" } });

            //return;

            SpreadsheetDocument spreadsheetDocument =
                    SpreadsheetDocument.Create("d:\\test1.xlsx", SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart and Workbook objects.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();

            // Create Worksheet and SheetData objects.
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add a Sheets object.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook
                .AppendChild<Sheets>(new Sheets());

            // Append the new worksheet named "mySheet" and associate it 
            // with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart
                    .GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "mySheet"
            };
            sheets.Append(sheet);

            // Get the sheetData cell table.
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            // Add a row to the cell table.
            Row row;
            row = new Row() { RowIndex = 1 };
            sheetData.Append(row);

            // Add the cell to the cell table at A1.
            Cell refCell = null;
            Cell newCell = new Cell() { CellReference = "A1" };
            row.InsertBefore(newCell, refCell);

            // Set the cell value to be a numeric value of 123.
            newCell.CellValue = new CellValue("123");
            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);

            // Close the document.
            spreadsheetDocument.Close();

            Console.WriteLine("All done. Press a key");
            Console.ReadKey();

        }

        public static void CreateExcel(string path, string sheetName, string[][] data)
        {
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookpart = spreadsheet.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                SharedStringTablePart shareStringPart;

                if (workbookpart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = workbookpart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = workbookpart.AddNewPart<SharedStringTablePart>();
                }
                shareStringPart.SharedStringTable = new SharedStringTable();
                shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text("50")));
                shareStringPart.SharedStringTable.Save();


                Sheets sheets = spreadsheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                if (sheetName != null && data.Count() > 0)
                {
                    Sheet sheet = new Sheet()
                    {
                        Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = sheetName
                    };
                    SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                    for (int i = 0; i < data.Count(); i++)
                    {
                        sheetData.Append(CreateContentRow(data[i], i + 1));
                    }
                    sheets.Append(sheet);
                }
                workbookpart.Workbook.Save();
                worksheetPart.Worksheet.Save();
                spreadsheet.Close();
            }
        }

        private static Row CreateContentRow(string[] cells, int rowIndex)
        {
            Row row = new Row
            {
                RowIndex = (UInt32)rowIndex
            };
            for (int i = 0; i < cells.Length; i++)
            {
                Cell dataCell = createTextCell(i + 1, rowIndex, cells[i]);
                row.AppendChild(dataCell);
            }
            return row;
        }

        private static Cell createTextCell(int columnIndex, int rowIndex, object cellValue)
        {
            Cell cell = new Cell();
            cell.DataType = CellValues.InlineString;
            InlineString inlineString = new InlineString();
            Text t = new Text();
            t.Text = cellValue == null ? string.Empty : cellValue.ToString();
            cell.CellValue = new CellValue(t.Text);
            inlineString.AppendChild(t);
            cell.AppendChild(inlineString);
            return cell;
        }
    }
}

