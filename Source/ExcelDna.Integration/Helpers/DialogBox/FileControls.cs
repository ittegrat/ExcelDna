
namespace ExcelDna.Integration.Helpers.DialogBox
{

  // 18 - Linked file list box
  public class FileListBox : DimmableControl, ITriggerControl
  {
    public bool IsTrigger { get; set; }
    public FileListBox(string id = null) : base(ControlType.LinkedFileListBox, id) { }
  }

  // 19 - Linked drive and directory box
  public class DriveDirBox : DimmableControl, ITriggerControl
  {
    public bool IsTrigger { get; set; }
    public DriveDirBox(string id = null) : base(ControlType.LinkedDriveDirBox, id) { }
  }

  // 20 - Directory text box
  public class DirectoryTextBox : DimmableControl
  {
    public DirectoryTextBox(string id = null) : base(ControlType.DirectoryTextBox, id) { }
  }

}
