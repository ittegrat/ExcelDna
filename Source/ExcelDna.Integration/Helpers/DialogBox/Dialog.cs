using System;
using System.Collections.Generic;

namespace ExcelDna.Integration.Helpers.DialogBox
{

  public delegate bool TriggerHandler(Dialog dialog);

  public class Dialog : Control
  {

    readonly List<Control> controls = new List<Control>();
    readonly HashSet<string> ids = new HashSet<string>();
    int? selected;

    public string Title { get; set; }
    public Control TriggerControl { get; private set; }
    public string TriggerId => TriggerControl?.Id;

    public Dialog() : base(ControlType.Dialog, null) { }

    public Dialog Add(Control control) {
      if (control.Id != null && !ids.Add(control.Id))
        throw new ArgumentException("A control with the same Id has already been added.");
      controls.Add(control);
      return this;
    }
    public Dialog Add(OptionButton ob) {
      var i = controls.Count - 1;
      while (i > -1) {
        switch (controls[i]) {
          case OptionGroup og:
            og.Add(ob);
            return Add((Control)ob);
          case OptionButton _:
            --i; continue;
        }
        break;
      }
      throw new ArgumentException("An OptionButton must be preceded by an OptionGroup.");
    }
    public Dialog Add(LinkedListBox llb) {
      var i = controls.Count - 1;
      while (i > -1) {
        switch (controls[i]) {
          case StringEditBox _:
          case IntegerEditBox _:
          case NumberEditBox _:
            return Add((Control)llb);
          case StaticText _:
            --i; continue;
        }
        break;
      }
      throw new ArgumentException("A LinkedListBox must be preceded by an EditBox.");
    }
    public Dialog Add(FileListBox fb) {
      var i = controls.Count - 1;
      while (i > -1) {
        switch (controls[i]) {
          case TextEditBox _:
            return Add((Control)fb);
          case StaticText _:
            --i; continue;
        }
        break;
      }
      throw new ArgumentException("A FileListBox must be preceded by a TextEditBox.");
    }
    // A DriveDirBox without a FileListBox is theoretically possible (the control shows
    // files, subfolders and drives), but the use seems buggy.
    public Dialog Add(DriveDirBox db) {
      var i = controls.Count - 1;
      while (i > -1) {
        switch (controls[i]) {
          case FileListBox _:
            return Add((Control)db);
          case StaticText _:
            --i; continue;
        }
        break;
      }
      throw new ArgumentException("A DriveDirBox must be preceded by a FileListBox.");
    }
    public Dialog Add(LinkedDropDown ldd) {
      var i = controls.Count - 1;
      while (i > -1) {
        switch (controls[i]) {
          case StringEditBox _:
          case IntegerEditBox _:
          case NumberEditBox _:
            return Add((Control)ldd);
          case StaticText _:
            --i; continue;
        }
        break;
      }
      throw new ArgumentException("A LinkedDropDown must be preceded by an EditBox.");
    }

    public Control GetControl(string id) {
      if (id != null) {
        id = id.Trim();
        if (id.Length > 0)
          return
            controls.Find(c => c.Id == id)
            ?? throw new KeyNotFoundException("A control with the given Id is not present.")
          ;
      }
      throw new ArgumentException("Invalid empty id.");
    }
    public T GetControl<T>(string id) where T : Control {
      return (T)GetControl(id);
    }

    public void Enable(string id) {
      GetControl<DimmableControl>(id).Enabled = true;
    }
    public void Disable(string id) {
      GetControl<DimmableControl>(id).Enabled = false;
    }
    public void SetFocus(string id) {
      if (id != null) {
        id = id.Trim();
        if (id.Length > 0) {
          var index = controls.FindIndex(c => c.Id == id);
          if (index < 0)
            throw new KeyNotFoundException("A control with the given Id is not present.");
          selected = index;
          return;
        }
      }
      throw new ArgumentException("Invalid empty id.");
    }
    public void ClearFocus() { selected = null; }

    public bool Show(TriggerHandler onTrigger = null) {

      bool handled;
      do {
        var ddef = BuildDef();
        var ans = XlCall.Excel(XlCall.xlfDialogBox, ddef);
        if (ans is bool b && b == false) {
          return false;
        }
        else {
          var result = ans as object[,];
          // xlfDialogBox always fill result[0, 6] with Max(1,selected trigger) so
          // pressing the enter key is distinguishable only if there is a non-trigger
          // control in the first position.
          selected = Convert.ToInt32(result[0, 6]) - 1;
          var sc = controls[selected.Value];
          TriggerControl = (
            sc is ButtonControl ||
            (sc is OptionButton ob && ob.OptionGroup.IsTrigger) ||
            (sc is ITriggerControl tc && tc.IsTrigger)
          )
            ? sc
            : null
          ;
          for (int i = 0; i < controls.Count; ++i) {
            if (controls[i] is ValueControl vc)
              vc.State = result[1 + i, 6];
          }
          handled = onTrigger?.Invoke(this) ?? true;
        }
      } while (!handled);

      return handled;

    }

    object[,] BuildDef() {

      if (controls.Count == 0)
        throw new InvalidOperationException("No controls added.");

      var def = new object[1 + controls.Count, 7];

      BuildDef(0, ref def);
      def[0, 5] = Title;
      if (selected.HasValue)
        def[0, 6] = 1 + selected.Value;

      for (var i = 0; i < controls.Count; ++i)
        controls[i].BuildDef(1 + i, ref def);

      return def;

    }

  }

  public static class DialogExtensions
  {
    public static T GetValue<T>(this Dialog dialog, string id) {
      var control = dialog.GetControl(id);
      switch (control) {
        case StringEditBox seb:
          return (T)Convert.ChangeType(seb.Value, typeof(T));
        case IntegerEditBox ieb:
          return (T)Convert.ChangeType(ieb.Value, typeof(T));
        case NumberEditBox neb:
          return (T)Convert.ChangeType(neb.Value, typeof(T));
        case CheckBox cb:
          return (T)Convert.ChangeType(cb.Value, typeof(T));
        case ListControl lc:
          return (T)Convert.ChangeType(lc.SelectedItem, typeof(T));
        default:
          throw new ArgumentException($"Control '{id}': unhandled type '{control.GetType().Name}'.");
      }
    }
    public static void SetValue(this Dialog dialog, string id, object value) {
      var control = dialog.GetControl(id);
      switch (control) {
        case StringEditBox seb:
          seb.Value = (string)value; return;
        case IntegerEditBox ieb:
          ieb.Value = (short?)value; return;
        case NumberEditBox neb:
          neb.Value = (double?)value; return;
        case CheckBox cb:
          cb.Value = (bool?)value; return;
        case ListControl lc:
          lc.SelectedItem = value; return;
        default:
          throw new ArgumentException($"Control '{id}': unhandled type '{control.GetType().Name}'.");
      }
    }
  }

}
