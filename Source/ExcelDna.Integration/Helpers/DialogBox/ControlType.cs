
namespace ExcelDna.Integration.Helpers.DialogBox
{
  public abstract partial class Control
  {

    protected enum ControlType
    {
      /// DialogBox (special, not valid as DialogBox item)
      Dialog = 0,
      /// Default OK button
      DefaultOkButton = 1,
      /// Cancel button
      CancelButton = 2,
      /// OK button
      OkButton = 3,
      /// Default Cancel button
      DefaultCancelButton = 4,
      /// Static text
      StaticText = 5,
      /// Text edit box
      TextEditBox = 6,
      /// Integer edit box
      IntegerEditBox = 7,
      /// Number edit box
      NumberEditBox = 8,
      /// Formula edit box
      FormulaEditBox = 9,
      /// Reference edit box
      ReferenceEditBox = 10,
      /// Option button group
      OptionButtonGroup = 11,
      /// Option button
      OptionButton = 12,
      /// Check box
      CheckBox = 13,
      /// Group box
      GroupBox = 14,
      /// List box
      ListBox = 15,
      /// Linked list box
      LinkedListBox = 16,
      /// Icons
      Icons = 17,
      /// Linked file list box (Microsoft Excel for Windows only)
      LinkedFileListBox = 18,
      /// Linked drive and directory box (Microsoft Excel for Windows only)
      LinkedDriveDirBox = 19,
      /// Directory text box
      DirectoryTextBox = 20,
      /// Drop-down list box
      DropDownListBox = 21,
      /// Drop-down combination edit/list box
      LinkedDropDown = 22,
      /// Picture button
      PictureButton = 23,
      /// Help button
      HelpButton = 24,
    }

  }
}
