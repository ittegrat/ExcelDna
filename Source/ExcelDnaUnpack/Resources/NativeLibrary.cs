
namespace ExcelDnaUnpack
{
  internal abstract partial class Resource
  {
    private sealed class NativeLibrary : Resource
    {

      public NativeLibrary(ResourceType rt, string type, string name, Module module)
        : base(rt, type, name, module) { }

    }
  }
}
