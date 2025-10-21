
namespace ExcelDnaUnpack
{
  internal abstract partial class Resource
  {
    private sealed class Doc : Resource
    {

      public Doc(ResourceType rt, string type, string name, Module module)
        : base(rt, type, name, module) { }

    }
  }
}
