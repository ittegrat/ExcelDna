
namespace ExcelDna.Integration.Helpers.DialogBox
{

  public abstract class TextControl : DimmableControl
  {

    public string Text { get; set; }

    protected TextControl(ControlType type, string id) : base(type, id) { }
    internal override void BuildDef(int row, ref object[,] def) {
      base.BuildDef(row, ref def);
      def[row, 5] = Text;
    }

  }

  public abstract class ButtonControl : TextControl
  {
    protected ButtonControl(ControlType type, string id) : base(type, id) { }
  }

  // 1 - Default Ok Button
  public class OkButtonDef : ButtonControl
  {
    public OkButtonDef(string id = null) : base(ControlType.DefaultOkButton, id) { Text = "OK"; }
  }

  // 2 - Cancel button
  public class CancelButton : ButtonControl
  {
    public CancelButton(string id = null) : base(ControlType.CancelButton, id) { Text = "Cancel"; }
  }

  // 3 - Ok button
  public class OkButton : ButtonControl
  {
    public OkButton(string id = null) : base(ControlType.OkButton, id) { }
  }

  // 4 - Default Cancel Button
  public class CancelButtonDef : ButtonControl
  {
    public CancelButtonDef(string id = null) : base(ControlType.DefaultCancelButton, id) { Text = "Cancel"; }
  }

  // 5 - Static text
  public class StaticText : TextControl
  {
    public StaticText(string id = null) : base(ControlType.StaticText, id) { }
  }

  // 14 - Group box
  public class GroupBox : TextControl
  {
    public GroupBox(string id = null) : base(ControlType.GroupBox, id) { }
  }

}
