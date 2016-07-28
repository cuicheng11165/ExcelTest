using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ExcelUtl
{
    public class SpreadsheetImport : IDisposable
    {
        SpreadsheetDocument spreadsheet;
        Sheets sheets;
        WorkbookPart workbookpart;
        WorksheetPart worksheetPart;
        Sheet sheet;
        uint rowIndex = 1;
        bool columnHeaderGerated = false;
        private Dictionary<PropertyInfo, string> properties;

        public SpreadsheetImport(string outPutFilePath, string sheetName)
        {
            // Create a spreadsheet document by supplying the file name.
            spreadsheet = SpreadsheetDocument.Create(outPutFilePath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            workbookpart = spreadsheet.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            sheets = spreadsheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            sheet = new Sheet()
            {
                Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = sheetName
            };

        }

        public void AddElements<T>(IList<T> elements)
        {
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            if (properties == null)
            {
                properties = GetProperties(typeof(T));
                sheetData.Append(CreateColumnHeader(properties));
            }

            sheetData.Append(elements.Select(CreateContentRow));

        }

        private Row CreateContentRow<T>(T element)
        {
            var row = new Row() { RowIndex = rowIndex++ };
            int cellIndex = 0;
            foreach (var entry in properties)
            {
                var value = entry.Key.GetValue(element, null);
                row.AppendChild(CreateCell(value, cellIndex++, entry.Value));
            }
            return row;
        }



        private Row CreateColumnHeader(Dictionary<PropertyInfo, string> properties)
        {
            var row = new Row() { RowIndex = rowIndex++ };
            int cellIndex = 0;
            if (!columnHeaderGerated)
            {
                foreach (var entry in properties)
                {
                    row.AppendChild(CreateCell(entry.Value, cellIndex++, entry.Value));
                }
                columnHeaderGerated = true;
            }
            return row;
        }

        private static Dictionary<PropertyInfo, string> GetProperties(Type type)
        {
            Dictionary<PropertyInfo, string> properties = new Dictionary<PropertyInfo, string>();

            foreach (var propertyInfo in type.GetProperties())
            {
                var attributes = propertyInfo.GetCustomAttributes(typeof(SheetColumn), true);

                string columnHeader = string.Empty;
                bool hasSheetColumnAttribute = false;

                foreach (var attr in attributes)
                {
                    var sheetColumnAttr = attr as SheetColumn;

                    if (sheetColumnAttr != null)
                    {
                        columnHeader = sheetColumnAttr.ColumnName;
                        hasSheetColumnAttribute = true;
                    }
                }

                if (!hasSheetColumnAttribute)
                {
                    continue;
                }

                if (columnHeader == string.Empty)
                {
                    columnHeader = propertyInfo.Name;
                }
                properties[propertyInfo] = columnHeader;
            }

            return properties;
        }

        private Cell CreateCell(object cellValue, int cellIndex, string headerName)
        {
            if (cellValue is Int32 || cellValue is double)
            {
                return new NumberCell(headerName, cellValue.ToString(), cellIndex);
            }
            else if (cellValue is bool)
            {
                return new BooleanCell(headerName, (bool)cellValue, cellIndex);
            }
            return new TextCell(headerName, cellValue == null ? string.Empty : cellValue.ToString(), cellIndex);
        }

        public void Dispose()
        {
            sheets.Append(sheet);
            workbookpart.Workbook.Save();
            worksheetPart.Worksheet.Save();
            spreadsheet.Close();
        }
    }

    public class SheetColumn : Attribute
    {
        public SheetColumn(string columnName) { this.ColumnName = columnName; }
        public string ColumnName { get; set; }
    }


    public class BooleanCell : Cell
    {
        public BooleanCell(string header, bool value, int index)
        {
            this.DataType = CellValues.Boolean;
            this.CellReference = header + index;
            //Add text to the text cell.
            this.CellValue = new CellValue(BooleanValue.FromBoolean(value));
        }
    }
    public class TextCell : Cell
    {
        public TextCell(string header, string text, int index)
        {
            this.DataType = CellValues.InlineString;
            this.CellReference = header + index;
            //Add text to the text cell.
            this.InlineString = new InlineString { Text = new Text { Text = text } };
        }
    }
    public class NumberCell : Cell
    {
        public NumberCell(string header, string text, int index)
        {
            this.DataType = CellValues.Number;
            this.CellReference = header + index;
            this.CellValue = new CellValue(text);
        }
    }

    public class HeaderCell : TextCell
    {
        public HeaderCell(string header, string text, int index) :
            base(header, text, index)
        {
            this.StyleIndex = 11;
        }
    }
}
