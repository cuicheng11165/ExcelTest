using System;

namespace Spreadsheet.Serialization
{
    public class TestEntity
    {
        [SheetColumnTest("Column1", true)]
        public string TestString { set; get; }

        [SheetColumnTest("Column2")]
        public int TestInt32 { set; get; }

        [SheetColumnTest("Column3")]
        public double TestDouble { set; get; }

        [SheetColumnTest("Column4")]
        public bool TestBoolean { set; get; }

        [SheetColumnTest("Column5")]
        public DateTime TestDateTime { set; get; }
    }

    public class SheetColumnTest : Attribute
    {
        public SheetColumnTest(string columnName) { this.ColumnName = columnName; }
        public SheetColumnTest(string columnName, bool hidden)
            : this(columnName)
        { this.Hidden = hidden; }
        string ColumnName { get; set; }
        public bool Hidden { set; get; }
    }
}