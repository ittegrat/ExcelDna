using System;
using SevenZip.Compression.LZMA;

namespace ExcelDnaUnpack
{
  internal abstract partial class Resource
  {
    private sealed class Assembly : Resource
    {

      public Assembly(ResourceType rt, string type, string name, Module module)
        : base(rt, type, name, module) { }
      public override string GetFileName() { return Name + ".dll"; }
      public bool IsEncoded() {
        var rbytes = module.GetResourceBytes(Type, Name);
        var bytes = IsCompressed()
          ? SevenZipHelper.Decompress(rbytes)
          : rbytes
        ;
        if (bytes[0] == 0x4d && bytes[1] == 0x5a)  // MZ Header
          return false;
        else if (bytes[0] == 0x08 && bytes[1] == 0x22)  // MZ Header XOR encoded
          return true;
        else
          throw new InvalidOperationException($"Packed files encoding can't be automatically detected\nUse -d to force behavior.");
      }

    }
  }
}
