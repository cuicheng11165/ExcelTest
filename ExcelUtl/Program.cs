using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Spreadsheet.Serialization
{
    class Program
    {
        static void Main(string[] args)
        {

            var entity = new TestEntity
            {
                TestString = "aqadsdasd",
                TestInt32 = 100,
                TestDouble = 1.2122,
                TestBoolean = true,
                TestDateTime = DateTime.Now
            };
            var entity1 = new TestEntity
            {
                TestString = @"sdfsdfasdfsadfsad
fdfsdfsdf
sfdsdfsdf
sfdsdfsf
sdfsf
sadfhoi",
                TestInt32 = 100,
                TestDouble = 3.1415926,
                TestBoolean = false,
                TestDateTime = DateTime.Now
            };


            var list = new List<TestEntity>() { entity, entity1 };

            var serializer = new SpreadsheetSerializer<TestEntity, SheetColumnTest>("ColumnName");

            serializer.ErrorNotify += (sender, convertArgs) =>
            {
                Console.WriteLine("Convert Failed when convert {0} to {1} . Row {2} , Column {3}, Error Message {4}.", convertArgs.Value, convertArgs.BindingType, convertArgs.RowIndex, convertArgs.ColumnIndex, convertArgs.ErrorException);
            };

            serializer.Serialize("d:\\test2.xlsx", list);

            var result1 = serializer.Deserialize("d:\\test2.xlsx");
            var result = serializer.Deserialize("d:\\Test.xlsx");



        }

        public static string ReadXml(string file)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(file, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;

                IEnumerable<Sheet> sheets = workbookPart.Workbook.Descendants<Sheet>();

                if (sheets.Count() == 0)
                {
                    return null;
                }

                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheets.First().Id);

                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                return sheetData.OuterXml;
            }
        }



    }
}

