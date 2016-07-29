using System;

namespace Spreadsheet.Serialization
{
    public class SheetColumnTest : Attribute
    {
        public SheetColumnTest(string columnName) { this.ColumnName = columnName; }
        string ColumnName { get; set; }
    }
}