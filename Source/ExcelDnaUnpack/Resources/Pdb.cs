
namespace ExcelDnaUnpack
{
  internal abstract partial class Resource
  {
    private sealed class Pdb : Resource
    {

      public Pdb(ResourceType rt, string type, string name, Module module)
        : base(rt, type, name, module) { }
      public override string GetFileName() { return Name + ".pdb"; }

    }
  }
}
