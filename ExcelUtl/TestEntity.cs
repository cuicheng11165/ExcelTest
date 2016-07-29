using System;

namespace Spreadsheet.Serialization
{
    public class TestEntity
    {
        [SheetColumnTest("")]
        public string TestString { set; get; }

        [SheetColumnTest("")]
        public int TestInt32 { set; get; }

        [SheetColumnTest("")]
        public double TestDouble { set; get; }

        [SheetColumnTest("")]
        public bool TestBoolean { set; get; }

        [SheetColumnTest("")]
        public DateTime TestDateTime { set; get; }
    }
}