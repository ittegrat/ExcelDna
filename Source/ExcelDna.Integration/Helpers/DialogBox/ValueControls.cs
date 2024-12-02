using System;
using System.Text;

namespace ExcelDna.Integration.Helpers.DialogBox
{

  public abstract class ValueControl : DimmableControl
  {
    internal object State { get; set; }
    protected ValueControl(ControlType type, string id) : base(type, id) { }
    internal override void BuildDef(int row, ref object[,] def) {
      base.BuildDef(row, ref def);
      def[row, 6] = State;
    }
    protected InvalidOperationException GetStateException() {
      var sb = new StringBuilder(type.ToString());
      if (Id != null) sb.Append(" '").Append(Id).Append("'");
      sb.Append(": invalid state value '");
      sb.Append(State is null ? "null" : State);
      sb.Append("'.");
      return new InvalidOperationException(sb.ToString());
    }
  }

  public abstract class StringEditBox : ValueControl
  {
    public string Value {
      get {
        switch (State) {
          case string s:
            return s;
          case null:
          case ExcelEmpty e:
            return null;
        }
        throw GetStateException();
      }
      set { State = value; }
    }
    protected StringEditBox(ControlType type, string id) : base(type, id) { }
  }

  // 6 - Text edit box
  public class TextEditBox : StringEditBox
  {
    public TextEditBox(string id = null) : base(ControlType.TextEditBox, id) { }
  }

  // 7 - Integer edit box
  public class IntegerEditBox : ValueControl
  {
    public short? Value {
      get {
        switch (State) {
          case double d:
            return (short)d;
          case null:
          case ExcelEmpty e:
            return null;
        }
        throw GetStateException();
      }
      set {
        if (value < -32766) throw new ArgumentOutOfRangeException(nameof(Value), value, $"The value {value} is not supported.");
        State = value;
      }
    }
    public IntegerEditBox(string id = null) : base(ControlType.IntegerEditBox, id) { }
  }

  // 8 - Number edit box
  public class NumberEditBox : ValueControl
  {
    public double? Value {
      get {
        switch (State) {
          case double d:
            return d;
          case null:
          case ExcelEmpty e:
            return null;
        }
        throw GetStateException();
      }
      set { State = value; }
    }
    public NumberEditBox(string id = null) : base(ControlType.NumberEditBox, id) { }
  }

  // 9 - Formula edit box
  public class FormulaEditBox : StringEditBox
  {
    public FormulaEditBox(string id = null) : base(ControlType.FormulaEditBox, id) { }
  }

  // 10 - Reference edit box
  public class RefEditBox : StringEditBox
  {
    public RefEditBox(string id = null) : base(ControlType.ReferenceEditBox, id) { }
  }

  // 13 - Check box
  public class CheckBox : ValueControl, ITriggerControl
  {

    public bool IsTrigger { get; set; }
    public string Text { get; set; }
    public bool? Value {
      get {
        switch (State) {
          case bool b:
            return b;
          case ExcelError e:
            if (e == ExcelError.ExcelErrorNA) return null;
            break;
          case null:
            return null;
        }
        throw GetStateException();
      }
      set {
        if (value.HasValue)
          State = value.Value;
        else
          State = ExcelError.ExcelErrorNA;
      }
    }

    public CheckBox(string id = null) : base(ControlType.CheckBox, id) { }
    internal override void BuildDef(int row, ref object[,] def) {
      base.BuildDef(row, ref def);
      def[row, 5] = Text;
    }

  }

}
