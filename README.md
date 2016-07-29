# Excel Serialization

     public class Student
     {
     	[ExcelColumn("Student Name")]
     	public string Name{get;set;}
     }
     
     public class ExcelColumn : Attribute
     {
    	public ExcelColumn(string columnName) { this.ColumnName = columnName; }
    	public string ColumnName { get; set; }
     }
    
## Export entity collection to excel
 
     var serializer = new SpreadsheetSerializer<Student, ExcelColumn>("ColumnName");
     serializer.Serialize("d:\\test2.xlsx", list);

## Import entity collection from excel

     var serializer = new SpreadsheetSerializer<Student, ExcelColumn>("ColumnName");
     serializer.ErrorNotify += (sender, convertArgs) =>
     {
     	Console.WriteLine("Convert Failed when convert {0} to {1} . Row {2} , Column {3}, Error Message {4}.", convertArgs.Value, convertArgs.BindingType, convertArgs.RowIndex, convertArgs.ColumnIndex, convertArgs.ErrorException);
     };
     var result = serializer.Deserialize("d:\\test2.xlsx");
