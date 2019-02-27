
namespace ExcelDnaUnpack
{
  internal abstract partial class Resource
  {
    private sealed class Assembly : Resource
    {

      public Assembly(ResourceType rt, string type, string name, Module module)
        : base(rt, type, name, module) { }
      public override string GetFileName() { return Name + ".dll"; }

    }
  }
}
