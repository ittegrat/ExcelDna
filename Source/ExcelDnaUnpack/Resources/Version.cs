using System;
using System.Diagnostics;
using System.Text;

namespace ExcelDnaUnpack
{
  internal abstract partial class Resource
  {
    private sealed class Version : Resource
    {

      readonly FileVersionInfo fvi = null;

      public bool IsFVI { get { return fvi != null; } }

      public override string BeautyType { get { return $"RT_VERSION ({Type})"; } }
      public override string BeautyName { get { return IsFVI ? $"VERSIONINFO ({Name})" : Name; } }

      public Version(ResourceType rt, string type, string name, Module module)
        : base(rt, type, name, module) {
        if (name == "#1")
          fvi = FileVersionInfo.GetVersionInfo(module.FileName);
      }
      public override byte[] GetBytes() {
        if (!IsFVI)
          throw new InvalidOperationException($"Unhandled VERSIONINFO resource '{Name}'.");
        var info = GetVersionInfo();
        var bytes = Encoding.UTF8.GetBytes(String.Join("\r\n", info));
        return bytes;
      }
      public override string GetFileName() { return "VersionInfo.txt"; }
      public string[] GetVersionInfo() {
        return new string[] {
          $"FileDescription: {fvi.FileDescription}",
          $"OriginalFilename: {fvi.OriginalFilename}",
          $"ProductVersion: {fvi.ProductVersion}",
          $"ProductVersion: {fvi.ProductMajorPart},{fvi.ProductMinorPart},{fvi.ProductBuildPart},{fvi.ProductPrivatePart}",
          $"FileVersion: {fvi.FileVersion}",
          $"FileVersion: {fvi.FileMajorPart},{fvi.FileMinorPart},{fvi.FileBuildPart},{fvi.FilePrivatePart}",
          $"IsDebug: {fvi.IsDebug}",
        };
      }

    }
  }
}
