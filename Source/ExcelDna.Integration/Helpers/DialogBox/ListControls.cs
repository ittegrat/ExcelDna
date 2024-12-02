using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDna.Integration.Helpers.DialogBox
{

  public abstract class ListControl : ValueControl, ITriggerControl
  {

    readonly List<object> items = new List<object>();
    string formula;

    public string Formula {
      get => formula;
      set {
        if (value != null) {
          value = value.Trim();
          if (value.Length == 0)
            throw new ArgumentException("Invalid empty formula.");
        }
        formula = value;
      }
    }
    public IReadOnlyList<object> Items => GetItems();
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
          if (Formula is null && (val < 0 || val >= items.Count))
            throw new ArgumentOutOfRangeException(
              nameof(SelectedIndex), val, $"{type}{(Id is null ? String.Empty : " '" + Id + "'")}: invalid value '{val}' for {nameof(SelectedIndex)}."
            );
          State = (double)(1 + val);
        }
        else
          State = ExcelError.ExcelErrorNA;
      }
    }
    public object SelectedItem {
      get { var items = GetItems();  return (SelectedIndex is int index) ? items[index] : null; }
      set { SelectedIndex = GetItems().IndexOf(value); }
    }

    public ListControl AddItem(object value) { GetItems().Add(value); return this; }
    public ListControl AddItemRange(IEnumerable<object> collection) { GetItems().AddRange(collection); return this; }
    public ListControl ClearItems() { GetItems().Clear(); SelectedIndex = null; return this; }
    public ListControl RemoveItems(int index, int count = 1) {
      GetItems().RemoveRange(index, count);
      if (SelectedIndex is int si) {
        var end = index + count;
        if (si >= index && si < end)
          SelectedIndex = null;
        else if (si >= end)
          SelectedIndex -= count;
      }
      return this;
    }

    internal override void BuildDef(int row, ref object[,] def) {
      string text;
      if (Formula is null) {
        // NB: a null, an empty string or an empty array crashes Excel.
        // Use { "" } for an empty list.
        var sb = new StringBuilder("{\"");
        for (var i = 0; i < items.Count; ++i) {
          if (i > 0) sb.Append("\",\"");
          sb.Append(items[i]);
        }
        sb.Append("\"}");
        text = sb.ToString();
      }
      else
        text = Formula;
      base.BuildDef(row, ref def);
      def[row, 5] = text;
    }

    protected ListControl(ControlType type, string id) : base(type, id) { }
    protected ListControl(ControlType type, IEnumerable<object> collection, string id = null) : base(type, id) { GetItems().AddRange(collection); }

    List<object> GetItems() {
      if (Formula != null)
        throw new InvalidOperationException($"{type}{(Id is null ? String.Empty : " '" + Id + "'")}: Items are not available when Formula is set.");
      return items;
    }

  }

  // 15 - List box
  public class ListBox : ListControl
  {
    public ListBox(string id = null) : base(ControlType.ListBox, id) { }
    public ListBox(IEnumerable<object> collection, string id = null) : base(ControlType.ListBox, collection, id) { }
  }

  // 16 - Linked list box
  public class LinkedListBox : ListControl
  {
    public LinkedListBox(string id = null) : base(ControlType.LinkedListBox, id) { }
    public LinkedListBox(IEnumerable<object> collection, string id = null) : base(ControlType.LinkedListBox, collection, id) { }
  }

  // 21 - Dropdown list box
  public class DropDown : ListControl
  {
    public DropDown(string id = null) : base(ControlType.DropDownListBox, id) { }
    public DropDown(IEnumerable<object> collection, string id = null) : base(ControlType.DropDownListBox, collection, id) { }
  }

  // 22 - Drop-down combination edit/list box
  public class LinkedDropDown : ListControl
  {
    public LinkedDropDown(string id = null) : base(ControlType.LinkedDropDown, id) { }
    public LinkedDropDown(IEnumerable<object> collection, string id = null) : base(ControlType.LinkedDropDown, collection, id) { }
  }

}
