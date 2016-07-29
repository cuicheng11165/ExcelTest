using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Spreadsheet.Serialization
{
    /// <summary>
    /// 
    /// public class Student
    /// {
    ///     [ExcelColumn("Student Name")]
    ///     public string Name{get;set;}
    /// }
    /// public class ExcelColumn : Attribute
    /// {
    ///    public ExcelColumn(string columnName) { this.ColumnName = columnName; }
    ///    public string ColumnName { get; set; }
    /// }
    /// 
    /// 此时调用方式为
    /// 
    /// var serializer = new SpreadsheetSerializer<Student, ExcelColumn>("ColumnName");
    /// serializer.Serialize("d:\\test2.xlsx", list);
    /// 
    /// serializer.ErrorNotify += (sender, convertArgs) =>
    /// {
    ///     Console.WriteLine("Convert Failed when convert {0} to {1} . Row {2} , Column {3}, Error Message {4}.", convertArgs.Value, convertArgs.BindingType, convertArgs.RowIndex, convertArgs.ColumnIndex, convertArgs.ErrorException);
    /// };
    /// var result = serializer.Deserialize("d:\\test2.xlsx");
    /// 
    /// </summary>
    /// <typeparam name="T">表示要序列化的Entity的类型,如上面的例子中T的类型为Student</typeparam>
    /// <typeparam name="TV">用来标识某个属性需要Import的标签类型，如上面的例子中TV的类型为ExcelColumn</typeparam>
    public class SpreadsheetSerializer<T, TV>
        where TV : class
        where T : new()
    {
        public string AttributeProperty { get; private set; }


        private Dictionary<PropertyInfo, string> propertyInfoNameMapping;

        private uint rowIndex = 1;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="attributeProperty">用来标识ColumnName的属性名</param>
        public SpreadsheetSerializer(string attributeProperty)
        {
            AttributeProperty = attributeProperty;
        }

        public void Serialize(string outPutFilePath, IList<T> elements)
        {
            var spreadsheet = SpreadsheetDocument.Create(outPutFilePath, SpreadsheetDocumentType.Workbook);
            InnerSerialize("sheet1", elements, spreadsheet);
        }

        public void Serialize(Stream stream, IList<T> elements)
        {
            var spreadsheet = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            InnerSerialize("sheet1", elements, spreadsheet);
        }

        private void InnerSerialize(string sheetName, IList<T> elements, SpreadsheetDocument spreadsheet)
        {
            // Add a WorkbookPart to the document.
            var workbookpart = spreadsheet.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();
            var stylesPart = spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>();

            CreateStyleSheet(stylesPart);

            // Add a WorksheetPart to the WorkbookPart.
            var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            var sheets = spreadsheet.WorkbookPart.Workbook.AppendChild(new Sheets());

            var sheet = new Sheet
            {
                Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = sheetName
            };

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            propertyInfoNameMapping = GetProperties();

            var headerRow = new Row { RowIndex = rowIndex++ };
            var cellIndex = 0;
            foreach (var entry in propertyInfoNameMapping)
            {
                headerRow.AppendChild(CreateCellByElement(entry.Value, cellIndex++, headerRow.RowIndex));
            }
            sheetData.Append(headerRow);

            sheetData.Append(elements.Select(CreateRowByCell));

            sheets.Append(sheet);
            workbookpart.Workbook.Save();
            worksheetPart.Worksheet.Save();
            spreadsheet.Close();
        }

        private static void CreateStyleSheet(WorkbookStylesPart stylesPart)
        {
            stylesPart.Stylesheet = new Stylesheet { Fonts = new DocumentFormat.OpenXml.Spreadsheet.Fonts { Count = 1 } };
            // blank font list
            stylesPart.Stylesheet.Fonts.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Font());

            // create fills
            stylesPart.Stylesheet.Fills = new Fills();

            //create a solid red fill
            var solidRed = new PatternFill
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FFFF0000") },
                BackgroundColor = new BackgroundColor { Indexed = 64 }
            };
            // red fill

            stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } });
            // required, reserved by Excel
            stylesPart.Stylesheet.Fills.AppendChild(new Fill
            {
                PatternFill = new PatternFill { PatternType = PatternValues.Gray125 }
            }); // required, reserved by Excel
            stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = solidRed });
            stylesPart.Stylesheet.Fills.Count = 3;

            //blank border list
            stylesPart.Stylesheet.Borders = new Borders { Count = 1 };
            stylesPart.Stylesheet.Borders.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Border());

            // blank cell format list
            stylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats { Count = 1 };
            stylesPart.Stylesheet.CellStyleFormats.AppendChild(new CellFormat());

            stylesPart.Stylesheet.NumberingFormats = new NumberingFormats();
            stylesPart.Stylesheet.NumberingFormats.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.NumberingFormat()
            {
                NumberFormatId = 165,
                FormatCode = @"yyyy/mm/dd h:mm:ss"
            });
            stylesPart.Stylesheet.NumberingFormats.Count = 1;

            // cell format list
            stylesPart.Stylesheet.CellFormats = new CellFormats();
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat());
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat
            {
                NumberFormatId = 165,
                ApplyNumberFormat = new BooleanValue(true)
            });
            stylesPart.Stylesheet.CellFormats.Count = 2;

            stylesPart.Stylesheet.Save();
        }


        private Row CreateRowByCell(T element)
        {
            var row = new Row { RowIndex = rowIndex++ };
            var cellIndex = 0;


            foreach (var entry in propertyInfoNameMapping)
            {
                var value = entry.Key.GetValue(element);
                row.AppendChild(CreateCellByElement(value, cellIndex++, row.RowIndex));
            }
            return row;
        }

        private Dictionary<PropertyInfo, string> GetProperties()
        {
            var properties = new Dictionary<PropertyInfo, string>();

            foreach (var propertyInfo in typeof(T).GetProperties())
            {
                var attributes = propertyInfo.GetCustomAttributes(typeof(TV), true);

                var columnHeader = string.Empty;
                var hasColumnAttribute = false;

                foreach (var attr in attributes)
                {
                    var sheetColumnAttr = attr as TV;

                    if (sheetColumnAttr != null)
                    {
                        var attributePropertyInfo = typeof(TV).GetProperty(this.AttributeProperty, BindingFlags.Instance | BindingFlags.NonPublic);
                        columnHeader = (string)attributePropertyInfo.GetValue(sheetColumnAttr);
                        hasColumnAttribute = true;
                    }
                }

                if (!hasColumnAttribute)
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

        /// <summary>
        /// 列名是个26进制的数
        /// </summary>
        /// <param name="cellIndex"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        private string CreateHeaderIndex(int cellIndex, uint rowIndex)
        {
            string result = string.Empty;

            var current = cellIndex;
            while (true)
            {
                var last = current % 26;
                char lastChar = (char)('A' + last);
                result = lastChar + result;
                current /= 26;
                if (current == 0)
                {
                    break;
                }
            }
            return result + rowIndex;
        }

        private Cell CreateCellByElement(object cellValue, int cellIndex, uint rowIndex)
        {
            if (cellValue is Int32 || cellValue is float || cellValue is Int16 || cellValue is Int64 || cellValue is double)
            {
                return new Cell
                {
                    CellReference = CreateHeaderIndex(cellIndex, rowIndex),
                    DataType = CellValues.Number,
                    CellValue = new CellValue(cellValue.ToString())
                };
            }
            if (cellValue is DateTime)
            {
                return new Cell
                {
                    DataType = CellValues.Number,//DateTime 是以Number形式存的
                    CellReference = CreateHeaderIndex(cellIndex, rowIndex),
                    CellValue = new CellValue(DateTime.Now.ToOADate().ToString()),
                    StyleIndex = 1
                };
            }
            return new Cell
            {
                DataType = CellValues.InlineString,
                CellReference = CreateHeaderIndex(cellIndex, rowIndex),
                InlineString = new InlineString
                {
                    Text = new DocumentFormat.OpenXml.Spreadsheet.Text
                    {
                        Text = cellValue == null ? string.Empty : cellValue.ToString()
                    }
                }
            };
        }

        public List<T> Deserialize(string inputFile)
        {
            using (var filestream = File.OpenRead(inputFile))
            {
                return Deserialize(filestream);
            }
        }

        public List<T> Deserialize(Stream stream)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;

                IEnumerable<Sheet> sheets = workbookPart.Workbook.Descendants<Sheet>();

                if (sheets.Count() == 0)
                {
                    return null;
                }

                List<T> results = new List<T>();
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheets.First().Id);

                SharedStringTable stringTable = null;
                if (workbookPart.SharedStringTablePart != null)
                {
                    stringTable = workbookPart.SharedStringTablePart.SharedStringTable;
                }

                var properties = GetPropertyNameInfoMapping();
                string[] columnSlot = null;
                foreach (Row row in worksheetPart.Worksheet.Descendants<Row>())
                {
                    if (row.RowIndex == 1)
                    {//header
                        columnSlot = GetColumnNames(row, stringTable);
                    }
                    else
                    {
                        results.Add(CreatEntityByRow(row, columnSlot, properties, stringTable));
                    }
                }
                return results;
            }
        }

        private string GetValue(Cell cell, SharedStringTable stringTable)
        {
            if (cell.DataType == null)
            {
                return cell.InnerText;
            }
            if (cell.DataType == CellValues.SharedString)
            {
                return stringTable.ChildElements[Int32.Parse(cell.CellValue.InnerText)].InnerText;
            }
            if (cell.DataType == CellValues.InlineString)
            {
                return cell.InlineString.InnerText;
            }
            return String.Empty;
        }

        private string[] GetColumnNames(Row row, SharedStringTable stringTable)
        {
            List<string> columnSlots = new List<string>();

            foreach (var child in row)
            {
                var cell = child as Cell;
                if (cell != null)
                {
                    string cellVal = GetValue(cell, stringTable);
                    if (String.IsNullOrEmpty(cellVal))
                    {
                        columnSlots.Add(string.Empty);
                    }
                    else if (cell.CellReference != null)
                    {
                        columnSlots.Add(cellVal);
                    }
                }
            }
            return columnSlots.ToArray();
        }

        private T CreatEntityByRow(Row row, string[] columnSlot, Dictionary<string, PropertyInfo> properties, SharedStringTable stringTable)
        {
            var result = Activator.CreateInstance<T>();

            for (int cellIndex = 0; cellIndex < row.ChildElements.Count; cellIndex++)
            {
                var cell = row.ChildElements[cellIndex] as Cell;
                if (cell != null)
                {
                    string columName = columnSlot[cellIndex];

                    if (properties.ContainsKey(columName))
                    {
                        var propertyInfo = properties[columName];
                        if (cell.DataType == null || cell.DataType == CellValues.Number || cell.DataType == CellValues.Date || cell.DataType == CellValues.Boolean)
                        {
                            SetValue(result, propertyInfo, cell.InnerText, row.RowIndex, cellIndex);
                        }
                        else if (cell.DataType == CellValues.InlineString)
                        {
                            SetValue(result, propertyInfo, cell.InlineString.InnerText, row.RowIndex, cellIndex);
                        }
                        else if (cell.DataType == CellValues.SharedString)
                        {
                            var stringValue = stringTable.ChildElements[Int32.Parse(cell.InnerText)].InnerText;
                            SetValue(result, propertyInfo, stringValue, row.RowIndex, cellIndex);
                        }
                    }
                }
            }
            return result;
        }

        private void SetValue(T result, PropertyInfo propertyInfo, string value, uint curRowIndex, int cellIndex)
        {
            try
            {
                if (propertyInfo.PropertyType == typeof(string))
                {
                    propertyInfo.SetValue(result, value);
                }
                else if (propertyInfo.PropertyType == typeof(int))
                {
                    propertyInfo.SetValue(result, Convert.ToInt32(value));
                }
                else if (propertyInfo.PropertyType == typeof(double))
                {
                    propertyInfo.SetValue(result, Convert.ToDouble(value));
                }
                else if (propertyInfo.PropertyType == typeof(bool))
                {
                    var boolValue = string.Equals(value, "TRUE", StringComparison.OrdinalIgnoreCase) ||
                                   string.Equals(value, "1", StringComparison.OrdinalIgnoreCase);

                    propertyInfo.SetValue(result, boolValue);
                }
                else if (propertyInfo.PropertyType == typeof(DateTime))
                {
                    double oaDateValue;
                    if (double.TryParse(value, out oaDateValue))
                    {
                        var dateValue = DateTime.FromOADate(oaDateValue);
                        propertyInfo.SetValue(result, dateValue);
                    }
                    else
                    {
                        propertyInfo.SetValue(result, Convert.ToDateTime(value));
                    }
                }
            }
            catch (Exception e)
            {
                if (ErrorNotify != null)
                {
                    ErrorNotify.Invoke(result, new SpreedSheetConvertArgs()
                    {
                        ColumnIndex = cellIndex + 1,
                        RowIndex = curRowIndex,
                        Value = value,
                        BindingType = propertyInfo.PropertyType,
                        ErrorException = e.ToString(),
                    });
                }
            }
        }

        private Dictionary<string, PropertyInfo> GetPropertyNameInfoMapping()
        {
            Dictionary<string, PropertyInfo> properties = new Dictionary<string, PropertyInfo>();
            BindingFlags bindingFlags = BindingFlags.Public | BindingFlags.Instance;
            typeof(T).GetProperties(bindingFlags).Where(a =>
            {
                var att = a.GetCustomAttribute(typeof(TV));
                if (att == null)
                {
                    return false;
                }
                var attribute = att as TV;

                var propertyInfo = typeof(TV).GetProperty(this.AttributeProperty, BindingFlags.Instance | BindingFlags.NonPublic);

                var value = propertyInfo.GetValue(attribute);
                if (value == null || value.ToString() == "")
                {
                    //默认用属性名
                    properties.Add(a.Name, a);
                }
                else
                {
                    properties.Add(value.ToString(), a);
                }
                return true;
            }).ToList();
            return properties;
        }

        public event EventHandler<SpreedSheetConvertArgs> ErrorNotify;

    }

    public class SpreedSheetConvertArgs : EventArgs
    {
        public uint RowIndex { set; get; }

        public int ColumnIndex { set; get; }

        public object Value { set; get; }

        public Type BindingType { set; get; }

        public string ErrorException { set; get; }
    }
}
