using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelDnaUnpack
{

  internal static partial class NativeMethods
  {

    [Flags]
    public enum LoadLibraryFlags : uint
    {
      DONT_RESOLVE_DLL_REFERENCES = 0x00000001,
      LOAD_IGNORE_CODE_AUTHZ_LEVEL = 0x00000010,
      LOAD_LIBRARY_AS_DATAFILE = 0x00000002,
      LOAD_LIBRARY_AS_DATAFILE_EXCLUSIVE = 0x00000040,
      LOAD_LIBRARY_AS_IMAGE_RESOURCE = 0x00000020,
      LOAD_WITH_ALTERED_SEARCH_PATH = 0x00000008
    }

    const string KERNEL32_DLL = "Kernel32.dll";

    [DllImport(KERNEL32_DLL, CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr BeginUpdateResource(string lpFileName, [MarshalAs(UnmanagedType.Bool)] bool bDeleteExistingResources);

    [DllImport(KERNEL32_DLL, CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EndUpdateResource(IntPtr hUpdate, [MarshalAs(UnmanagedType.Bool)] bool fDiscard);

    public delegate bool EnumResTypeProc(IntPtr hModule, IntPtr lpszType, IntPtr lParam);
    [DllImport(KERNEL32_DLL, CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumResourceTypes(SafeModuleHandle hModule, EnumResTypeProc lpEnumFunc, IntPtr lParam);

    public delegate bool EnumResNameProc(IntPtr hModule, IntPtr lpszType, IntPtr lpszName, IntPtr lParam);
    [DllImport(KERNEL32_DLL, CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumResourceNames(SafeModuleHandle hModule, string lpType, EnumResNameProc lpEnumFunc, IntPtr lParam);

    [DllImport(KERNEL32_DLL, CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr FindResource(SafeModuleHandle hModule, string lpName, string lpType);

    [DllImport(KERNEL32_DLL, CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern SafeModuleHandle LoadLibraryEx(string lpFileName, IntPtr hFile, LoadLibraryFlags dwFlags);

    [DllImport(KERNEL32_DLL, SetLastError = true)]
    public static extern IntPtr LoadResource(SafeModuleHandle hModule, IntPtr hResource);

    [DllImport(KERNEL32_DLL)]
    public static extern IntPtr LockResource(IntPtr hResData);

    [DllImport(KERNEL32_DLL, SetLastError = true)]
    public static extern uint SizeofResource(SafeModuleHandle hModule, IntPtr hResource);

    [DllImport(KERNEL32_DLL, CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool UpdateResource(IntPtr hUpdate, string lpType, string lpName, ushort wLanguage, IntPtr lpData, uint cb);

  }

  internal sealed class Module : IDisposable
  {

    readonly Dictionary<ResourceType, List<Resource>> resources = new Dictionary<ResourceType, List<Resource>>();
    readonly SafeModuleHandle safeModuleHandle;
    readonly Version dnaVersion;

    public bool Decode { get; private set; }
    public string FileName { get; }

    public Module(string fName) {

      FileName = Path.GetFullPath(fName);

      if (!File.Exists(FileName))
        throw new ArgumentException($"Invalid xll file path: {FileName}");

      safeModuleHandle = LoadModule(FileName);

      var ans = NativeMethods.EnumResourceTypes(
        safeModuleHandle,
        AddResourceType,
        IntPtr.Zero
      );

      if (!ans) throw new Win32Exception();

      foreach (var key in resources.Keys) {
        var rl = resources[key];
        if (rl.Count > 0)
          rl.Sort();
        else
          resources.Remove(key);
      }

      if (resources.TryGetValue(ResourceType.VERSION, out var vl)) {
        if (Resource.GetVersion(vl) is Version ver) {
          dnaVersion = ver;
        }
      }

    }

    public void Dispose() {
      safeModuleHandle.Dispose();
    }
    public void Clean(string basePath, bool overwrite) {

      if (basePath == null)
        basePath = Path.GetDirectoryName(FileName);

      var outFileName = Path.Combine(
        basePath,
        Path.ChangeExtension(Path.GetFileName(FileName), ".Cleaned.xll")
      );

      if (!overwrite && File.Exists(outFileName))
        throw new ApplicationException($"Cleaned file '{outFileName}' exist.\nUse -f to force overwriting.");

      Console.WriteLine("Cleaning resources:\n");
      Console.WriteLine($"  Addin: {FileName}");
      Console.WriteLine($"  Clean: {outFileName}");
      Console.WriteLine();

      Directory.CreateDirectory(basePath);
      File.Copy(FileName, outFileName, true);

      IntPtr h = NativeMethods.BeginUpdateResource(outFileName, false);
      if (h == IntPtr.Zero) throw new Win32Exception();

      bool ans = false;
      try {

        foreach (var rl in resources.Values) {
          foreach (var r in rl) {

            if (r.ResourceType == ResourceType.UNKNOWN)
              continue;

            if (r.ResourceType == ResourceType.VERSION)
              continue;

            if (r.Name == "EXCELDNA.LOADER" || r.Name == "EXCELDNA.INTEGRATION")
              continue;

            Console.WriteLine($"  - {r.Type} - {r.Name}");

            ans = NativeMethods.UpdateResource(h, r.Type, r.Name, 0, IntPtr.Zero, 0);
            if (!ans) throw new Win32Exception();

          }
        }

      }
      catch (Exception ex) {
        Console.WriteLine(ex.Message);
      }

      ans = NativeMethods.EndUpdateResource(h, !ans);
      if (!ans) throw new Win32Exception();

    }
    public void Extract(ResourceType rt, bool? decode, string basePath, bool overwrite, bool mksubfolders) {

      if (decode.HasValue)
        Decode = decode.Value;
      else
        Decode = IsEncoded();

      if (basePath == null)
        basePath = Path.ChangeExtension(FileName, null);

      if (basePath == FileName)
        throw new ApplicationException($"Invalid output folder '{basePath}'.\nUse -o to set the output folder.");

      if (!overwrite && Directory.Exists(basePath))
        throw new ApplicationException($"Output folder '{basePath}' exist.\nUse -f to force extraction.");

      Console.WriteLine("Extracting resources:\n");
      Console.WriteLine($"  Addin: {FileName}");
      Console.WriteLine($"  OutPath: {basePath}");
      if (dnaVersion != null)
        Console.WriteLine($"  ExcelDna: v{dnaVersion}");
      Console.WriteLine($"  XorRecode: {(Decode ? "enabled" : "disabled")}");
      Console.WriteLine();

      foreach (var kv in resources) {

        //** MAYBE: how to extract unknown resources ?
        if (kv.Key == ResourceType.UNKNOWN)
          continue;

        if (rt.HasFlag(kv.Key)) {

          string outPath = basePath;
          if (mksubfolders) {
            outPath = Path.Combine(basePath, Resource.GetSubFolder(kv.Key));
          }

          try {
            ExtractResources(kv.Value, outPath);
          }
          catch (Exception ex) {
            Console.WriteLine(ex.Message);
          }

        }
      }

    }
    public void List(ResourceType rt) {

      Console.WriteLine("Listing resources:\n");
      Console.WriteLine($"  Addin: {FileName}");
      if (dnaVersion != null)
        Console.WriteLine($"  ExcelDna: v{dnaVersion}");

      foreach (var kv in resources) {
        if (rt.HasFlag(kv.Key))
          PrintResources(kv.Value);
      }

    }
    public void PrintVersion() {

      Console.WriteLine("Reading version info:\n");
      Console.WriteLine($"  Addin: {FileName}\n");

      if (!resources.TryGetValue(ResourceType.VERSION, out var rl)) {
        Console.WriteLine($"  [ERROR] No VERSIONINFO resource found.");
        return;
      }

      var lines = Resource.GetVersionInfo(rl);
      foreach (var line in lines) {
        Console.WriteLine($"  {line}");
      }

    }
    public byte[] GetResourceBytes(string type, string name) {

      var rh = NativeMethods.FindResource(safeModuleHandle, name, type);
      if (rh == IntPtr.Zero) throw new Win32Exception();

      var ph = NativeMethods.LoadResource(safeModuleHandle, rh);
      if (ph == IntPtr.Zero) throw new Win32Exception();

      var ptr = NativeMethods.LockResource(ph);
      if (ptr == IntPtr.Zero)
        throw new ApplicationException($"Invalid pointer from LockResource");

      var sz = NativeMethods.SizeofResource(safeModuleHandle, rh);
      if (sz == 0) throw new Win32Exception();

      var bytes = new byte[sz];
      Marshal.Copy(ptr, bytes, 0, (int)sz);

      return bytes;

    }

    void ExtractResources(List<Resource> resList, string outPath) {

      Directory.CreateDirectory(outPath);

      foreach (var r in resList) {

        var outFile = Path.Combine(
          outPath,
          r.GetFileName()
        );

        try {

          var bytes = r.GetBytes();

          Console.WriteLine($"  + {r.BeautyType} - {r.BeautyName}");

          using (var fs = new FileStream(outFile, FileMode.Create)) {
            fs.Write(bytes, 0, bytes.Length);
          }

        }
        catch (Exception ex) {
          Console.WriteLine(ex.Message);
        }

      }

    }
    void PrintResources(List<Resource> resList) {
      string resType = null;
      foreach (var r in resList) {
        if (!r.Type.Equals(resType)) {
          resType = r.Type;
          Console.WriteLine($"\n  {r.BeautyType}");
        }
        Console.WriteLine($"    {r.BeautyName}");
      }
    }
    bool IsEncoded() {
      if (resources.TryGetValue(ResourceType.ASSEMBLY, out var rl)) {
        var nc = rl.Find(r => !r.IsCompressed());
        return Resource.IsEncoded(nc ?? rl[0]);
      }
      throw new InvalidOperationException($"No ASSEMBLY resource packed. Encoding can't be automatically detected\nUse -d to force decoding.");
    }

    SafeModuleHandle LoadModule(string fileName) {

      var safeModuleHandle = NativeMethods.LoadLibraryEx(
        fileName,
        IntPtr.Zero,
        NativeMethods.LoadLibraryFlags.LOAD_LIBRARY_AS_DATAFILE | NativeMethods.LoadLibraryFlags.LOAD_LIBRARY_AS_IMAGE_RESOURCE
      );

      if (safeModuleHandle.IsInvalid) {
        var err = Marshal.GetLastWin32Error();
        switch (err) {
          case 193:
            throw new Win32Exception(err, $"{Path.GetFileName(fileName)} is not a valid Win32 application.");
          default:
            throw new Win32Exception(err);
        }
      }

      return safeModuleHandle;

    }
    bool AddResourceType(IntPtr hModule, IntPtr type, IntPtr lParam) {

      var typeName = GetResourceName(type);
      var rt = typeName.ToResourceType();

      if (!resources.ContainsKey(rt))
        resources.Add(rt, new List<Resource>());

      var ans = NativeMethods.EnumResourceNames(
        safeModuleHandle,
        typeName,
        AddResourceName,
        IntPtr.Zero
      );

      if (!ans) throw new Win32Exception();

      return true;

    }
    bool AddResourceName(IntPtr hModule, IntPtr type, IntPtr name, IntPtr lParam) {

      var typeName = GetResourceName(type);
      var rt = typeName.ToResourceType();

      if (!resources.TryGetValue(rt, out List<Resource> resList))
        return false;

      var resName = GetResourceName(name);
      resList.Add(Resource.Create(rt, this, typeName, resName));

      return true;

    }

    static bool IsIntResource(IntPtr value) {
      return (((ulong)value >> 16) == 0);
    }
    static string GetResourceName(IntPtr value) {
      if (IsIntResource(value)) {
        return $"#{value}";
      }
      return Marshal.PtrToStringUni(value);
    }

  }

}
