
namespace ExcelDnaUnpack
{
  internal abstract partial class Resource
  {
    private sealed class Dna : Resource
    {

      public Dna(ResourceType rt, string type, string name, Module module)
        : base(rt, type, name, module) { }
      public override string GetFileName() {
        if (Name == "__MAIN__") return Name + ".dna";
        return Name;
      }

    }
  }
}
