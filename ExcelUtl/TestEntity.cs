using System;

namespace Spreadsheet.Serialization
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
}