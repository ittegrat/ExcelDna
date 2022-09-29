using System;
using System.Collections.Generic;
using System.Linq;
using SevenZip.Compression.LZMA;

namespace ExcelDnaUnpack
{
  internal abstract partial class Resource : IComparable<Resource>
  {

    readonly Module module;

    public static Resource Create(ResourceType rt, Module module, string type, string name) {

      var baseType = typeof(Resource);

      var resType = baseType.Assembly.GetTypes()
        .Where(t =>
          baseType.IsAssignableFrom(t) &&
          0 == String.Compare(t.Name, rt.ToString(), StringComparison.InvariantCultureIgnoreCase)
        )
        .FirstOrDefault()
      ;

      if (resType != null)
        return (Resource)Activator.CreateInstance(resType, new object[] { rt, type, name, module });
      else
        return new Unknown(rt, type, name, module);

    }
    public static string GetSubFolder(ResourceType rt) {
      return rt.ToString();
    }
    public static System.Version GetVersion(List<Resource> resList) {
      var fvi = resList.Find(r => ((Version)r).IsFVI) as Version;
      return fvi?.GetVersion();
    }
    public static string[] GetVersionInfo(List<Resource> resList) {
      if (!(resList.Find(r => ((Version)r).IsFVI) is Version fvi))
        return new string[] { "[ERROR] No suitable VERSIONINFO resource found." };
      return fvi.GetVersionInfo();
    }

    public string Type { get; }
    public string Name { get; }
    public ResourceType ResourceType { get; }

    public virtual string BeautyType { get { return Type; } }
    public virtual string BeautyName { get { return Name; } }

    public int CompareTo(Resource other) {

      if (Type[0] == '#') {

        if (other.Type[0] == '#') {

          int t1 = int.Parse(Type.Substring(1));
          int t2 = int.Parse(other.Type.Substring(1));

          if (t1 == t2)
            return (int.Parse(Name.Substring(1)) - int.Parse(other.Name.Substring(1)));
          else
            return t1 - t2;

        } else
          return 1;

      }

      if (other.Type[0] == '#')
        return other.CompareTo(this);

      int t = String.Compare(Type, other.Type, StringComparison.InvariantCultureIgnoreCase);
      return (t != 0 ? t : String.Compare(Name, other.Name, StringComparison.InvariantCultureIgnoreCase));

    }

    public virtual byte[] GetBytes() {

      var rbytes = module.GetResourceBytes(Type, Name);

      var bytes = IsCompressed()
        ? SevenZipHelper.Decompress(rbytes)
        : rbytes
      ;

      if (module.Decode)
        XorRecode(bytes);

      return bytes;

    }
    public virtual string GetFileName() { return Name; }

    protected virtual bool IsCompressed() {
      return Type.EndsWith("_LZMA");
    }

    Resource(ResourceType rt, string type, string name, Module module) {
      ResourceType = rt;
      this.module = module;
      Type = type;
      Name = name;
    }

    static readonly byte[] xorKeys = System.Text.Encoding.ASCII.GetBytes("ExcelDna");
    static void XorRecode(byte[] data) {
      var keys = xorKeys;
      for (int i = 0; i < data.Length; i++) {
        data[i] = (byte)(keys[i % keys.Length] ^ data[i]);
      }
    }

  }
}
