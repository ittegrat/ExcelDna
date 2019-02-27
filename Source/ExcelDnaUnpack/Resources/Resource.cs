using System;
using System.Linq;
using SevenZip.Compression.LZMA;

namespace ExcelDnaUnpack
{
  internal abstract partial class Resource : IComparable<Resource>
  {

    Module module;

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

    public string Type { get; }
    public string Name { get; }
    public ResourceType ResourceType { get; }

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

      var bytes = module.GetResourceBytes(Type, Name);

      return IsCompressed()
        ? SevenZipHelper.Decompress(bytes)
        : bytes
      ;

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

  }
}
