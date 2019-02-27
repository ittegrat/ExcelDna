
namespace ExcelDnaUnpack
{
  internal abstract partial class Resource
  {
    private sealed class Unknown : Resource
    {

      public Unknown(ResourceType rt, string type, string name, Module module)
        : base(rt, type, name, module) { }

    }
  }
}
