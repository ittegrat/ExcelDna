using System;

namespace ExcelDna.Integration.Helpers.DialogBox
{

  public interface ITriggerControl
  {
    bool IsTrigger { get; set; }
  }

  public abstract partial class Control
  {

    protected readonly ControlType type;

    public string Id { get; }

    public int? Left { get; set; }
    public int? Top { get; set; }
    public int? Width { get; set; }
    public int? Height { get; set; }

    internal virtual void BuildDef(int row, ref object[,] def) {
      if (type != ControlType.Dialog) {
        var itemType = (int)type;
        if (this is DimmableControl dc) {
          if (dc.Enabled && this is ITriggerControl tr && tr.IsTrigger) itemType += 100;
          if (!dc.Enabled) itemType += 200;
        }
        def[row, 0] = itemType;
      }
      if (Left.HasValue) def[row, 1] = Left.Value;
      if (Top.HasValue) def[row, 2] = Top.Value;
      if (Width.HasValue) def[row, 3] = Width.Value;
      if (Height.HasValue) def[row, 4] = Height.Value;
    }

    protected Control(ControlType ct, string id) {
      if (id != null) {
        id = id.Trim();
        if (id.Length == 0)
          throw new ArgumentException("Invalid empty id.");
        Id = id;
      }
      type = ct;
    }

  }

  public abstract class DimmableControl : Control
  {
    public bool Enabled { get; set; } = true;
    protected DimmableControl(ControlType type, string id) : base(type, id) { }
  }

  // 17 - Icons
  public class Icon : Control
  {
    public enum Style : short { Question = 1, Information, Warning }
    public Style IconStyle { get; set; }
    public Icon(Style style, string id = null) : base(ControlType.Icons, id) { IconStyle = style; }
    internal override void BuildDef(int row, ref object[,] def) {
      base.BuildDef(row, ref def);
      def[row, 5] = (short)IconStyle;
    }
  }

}
