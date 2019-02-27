
namespace ExcelDnaUnpack
{
  internal abstract partial class Resource
  {
    private sealed class Image : Resource
    {

      public Image(ResourceType rt, string type, string name, Module module)
        : base(rt, type, name, module) { }

    }
  }
}
