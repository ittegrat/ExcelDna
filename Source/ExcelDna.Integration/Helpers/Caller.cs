
namespace ExcelDna.Integration.Helpers
{

  /*
   * Returns information about the cell, range of cells, command on a menu, tool on a toolbar,
   * or object that called the macro that is currently running. Use CALLER in a subroutine or
   * custom function whose behavior depends on the location, size, name, or other attribute
   * of the caller.
   *
   * 1. If the custom function is entered in a single cell, CALLER returns the reference of that cell.
   * 2. If the custom function was part of an array formula entered in a range of cells, CALLER returns
   *    the reference of the range.
   * 3. If CALLER appears in a macro called by an Auto_Open, Auto_Close, Auto_Activate, or Auto_Deactivate
   *    macro, it returns the name of the calling sheet.
   * 4. If CALLER appears in a macro called by an assigned-to-object macro, it returns the object identifier.
   * 5. If CALLER appears in a macro called by a command on a menu, it returns a horizontal array of
   *    three elements including the command's position number, the menu number, and the menu bar number.
   * 6. If CALLER appears in a macro called by a tool on a toolbar, it returns a horizontal array containing
   *    the position number and the toolbar name.
   * 7. If CALLER appears in a macro called by an ON.DOUBLECLICK or ON.ENTRY function, CALLER returns the
   *    name of the chart object identifier or cell reference, if applicable, to which the ON.DOUBLECLICK
   *    or ON.ENTRY macro applies.
   * 8. If CALLER appears in a macro that was run manually, or for any reason not described above, it
   *    returns the #REF! error value.
  */
  public class Caller
  {

    public enum ReturnType
    {
      Range,
      String,
      RefError,
      Other
    }

    public readonly ReturnType Type;

    public readonly ExcelReference XlReference;
    public readonly int Rows = 0;
    public readonly int Columns = 0;

    public readonly string Label;

    public Caller() {

      object obj = XlCall.Excel(XlCall.xlfCaller);

      if (obj is ExcelReference) {
        Type = ReturnType.Range;
        XlReference = (ExcelReference)obj;
        Rows = XlReference.RowLast - XlReference.RowFirst + 1;
        Columns = XlReference.ColumnLast - XlReference.ColumnFirst + 1;
      }
      else if (obj is ExcelError && (ExcelError)obj == ExcelError.ExcelErrorRef) {
        Type = ReturnType.RefError;
      }
      else if (obj is string) {
        Type = ReturnType.String;
        Label = (string)obj;
      }
      else
        Type = ReturnType.Other;

    }

    public bool IsRange { get { return Type == ReturnType.Range; } }
    public bool IsString { get { return Type == ReturnType.String; } }
    public bool IsRefError { get { return Type == ReturnType.RefError; } }

    public bool LackRows(int desiredRows) {
      return ((Type == ReturnType.Range) && (desiredRows>Rows));
    }
    public bool LackColumns(int desiredCols) {
      return ((Type == ReturnType.Range) && (desiredCols > Columns));
    }

  }

}
