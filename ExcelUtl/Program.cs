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
                TestDateTime = DateTime.MinValue
            };


            var list = new List<TestEntity>() { entity, entity1 };

            var serializer = new SpreadsheetSerializer<TestEntity, SheetColumnTest>("ColumnName");

            serializer.Serialize("d:\\test2.xlsx", list);

            var result = serializer.Deserialize("d:\\test2.xlsx");

        }

    }
}

