using System;

namespace Spreadsheet.Serialization
{
    public class SheetColumn : Attribute
    {
        public SheetColumn(string columnName) { this.ColumnName = columnName; }
        public string ColumnName { get; set; }
    }
}