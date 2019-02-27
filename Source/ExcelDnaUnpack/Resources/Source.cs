
namespace ExcelDnaUnpack
{
  internal abstract partial class Resource
  {
    private sealed class Source : Resource
    {

      public Source(ResourceType rt, string type, string name, Module module)
        : base(rt, type, name, module) { }

    }
  }
}
