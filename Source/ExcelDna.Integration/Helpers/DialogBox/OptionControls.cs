using System;
using System.Collections.Generic;

namespace ExcelDna.Integration.Helpers.DialogBox
{

  // 11 - Option button group
  public class OptionGroup : ValueControl, ITriggerControl
  {

    readonly List<OptionButton> options = new List<OptionButton>();

    public bool IsTrigger { get; set; }
    public int? SelectedIndex {
      get {
        switch (State) {
          case double d:
            return (int)d - 1;
          case ExcelError e:
            if (e == ExcelError.ExcelErrorNA) return null;
            break;
          case null:
            return null;
        }
        throw GetStateException();
      }
      set {
        if (value.HasValue) {
          var val = value.Value;
          if (val < 0 || val >= options.Count)
            throw new ArgumentOutOfRangeException(
              nameof(SelectedIndex), val, $"OptionGroup{(Id is null ? String.Empty : " '" + Id + "'")}: invalid value '{val}' for {nameof(SelectedIndex)}."
            );
          State = (double)(1 + val);
        }
        else
          State = ExcelError.ExcelErrorNA;
      }
    }
    public string SelectedId {
      get => SelectedItem?.Id;
      set {
        if (value != null) {
          value = value.Trim();
          if (value.Length == 0)
            throw new ArgumentException("Invalid empty id.");
          var index = options.FindIndex(c => c.Id == value);
          if (index < 0)
            throw new ArgumentOutOfRangeException(
              nameof(SelectedId), value, $"OptionGroup{(Id is null ? String.Empty : " '" + Id + "'")}: id '{value}' not found."
            );
          SelectedIndex = index;
        }
        else
          SelectedIndex = null;
      }
    }
    public OptionButton SelectedItem {
      get { return (SelectedIndex is int index) ? options[index] : null; }
    }
    public string Text { get; set; }

    public OptionGroup(string id = null) : base(ControlType.OptionButtonGroup, id) { }
    internal void Add(OptionButton button) { button.OptionGroup = this; options.Add(button); }
    internal override void BuildDef(int row, ref object[,] def) {
      if (options.Count == 0)
        throw new InvalidOperationException($"OptionGroup{(Id is null ? String.Empty : " '" + Id + "'")} without option buttons.");
      base.BuildDef(row, ref def);
      def[row, 5] = Text;
    }

  }

  // 12 - Option button
  public class OptionButton : TextControl, ITriggerControl
  {
    public bool IsTrigger { get; set; }
    public OptionGroup OptionGroup { get; internal set; }
    public OptionButton(string id = null) : base(ControlType.OptionButton, id) { }
  }

}
