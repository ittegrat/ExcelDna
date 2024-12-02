
namespace ExcelDna.Integration.Helpers.DialogBox
{
  public abstract partial class Control
  {

    protected enum ControlType
    {
      Dialog = 0,              /// DialogBox (special, not valid as DialogBox item)
      DefaultOkButton = 1,     /// Default OK button
      CancelButton = 2,        /// Cancel button
      OkButton = 3,            /// OK button
      DefaultCancelButton = 4, /// Default Cancel button
      StaticText = 5,          /// Static text
      TextEditBox = 6,         /// Text edit box
      IntegerEditBox = 7,      /// Integer edit box
      NumberEditBox = 8,       /// Number edit box
      FormulaEditBox = 9,      /// Formula edit box
      ReferenceEditBox = 10,   /// Reference edit box
      OptionButtonGroup = 11,  /// Option button group
      OptionButton = 12,       /// Option button
      CheckBox = 13,           /// Check box
      GroupBox = 14,           /// Group box
      ListBox = 15,            /// List box
      LinkedListBox = 16,      /// Linked list box
      Icons = 17,              /// Icons
      LinkedFileListBox = 18,  /// Linked file list box (Microsoft Excel for Windows only)
      LinkedDriveDirBox = 19,  /// Linked drive and directory box (Microsoft Excel for Windows only)
      DirectoryTextBox = 20,   /// Directory text box
      DropDownListBox = 21,    /// Drop-down list box
      LinkedDropDown = 22,     /// Drop-down combination edit/list box
      PictureButton = 23,      /// Picture button
      HelpButton = 24,         /// Help button
    }

  }
}
