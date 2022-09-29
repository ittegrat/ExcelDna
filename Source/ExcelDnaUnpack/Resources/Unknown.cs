using System;

namespace ExcelDnaUnpack
{
  internal abstract partial class Resource
  {
    private sealed class Unknown : Resource
    {

      public override string BeautyType { get { return Type + (Type == "#6" ? " - RT_STRING" : String.Empty); } }

      public Unknown(ResourceType rt, string type, string name, Module module)
        : base(rt, type, name, module) { }

    }
  }
}
