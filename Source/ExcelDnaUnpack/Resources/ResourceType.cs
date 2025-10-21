using System;

namespace ExcelDnaUnpack
{

  [Flags]
  internal enum ResourceType
  {
    NONE = 0x00,
    UNKNOWN = 0x01,
    ASSEMBLY = 0x02,
    CONFIG = 0x04,
    DNA = 0x08,
    IMAGE = 0x10,
    PDB = 0x20,
    SOURCE = 0x40,
    TYPELIB = 0x80,
    VERSION = 0x100,
    NATIVELIBRARY = 0x200,
    DOC = 0x400,
    KNOWN = ASSEMBLY | CONFIG | DNA | IMAGE | PDB | SOURCE | TYPELIB | VERSION | NATIVELIBRARY | DOC,
    ALL = KNOWN | UNKNOWN
  }

  internal static class ResourceTypeExtensions
  {

    public static ResourceType ToResourceType(this string typeName) {

      switch (typeName) {

        case "ASSEMBLY":
        case "ASSEMBLY_LZMA":
          return ResourceType.ASSEMBLY;

        case "CONFIG":
          return ResourceType.CONFIG;

        case "DNA":
        case "DNA_LZMA":
          return ResourceType.DNA;

        case "DOC":
        case "DOC_LZMA":
          return ResourceType.DOC;

        case "IMAGE":
        case "IMAGE_LZMA":
          return ResourceType.IMAGE;

        case "NATIVE_LIBRARY":
        case "NATIVE_LIBRARY_LZMA":
          return ResourceType.NATIVELIBRARY;

        case "PDB":
        case "PDB_LZMA":
          return ResourceType.PDB;

        case "SOURCE":
        case "SOURCE_LZMA":
          return ResourceType.SOURCE;

        case "TYPELIB":
          return ResourceType.TYPELIB;

        case "#16":
          return ResourceType.VERSION;

        default:
          return ResourceType.UNKNOWN;

      }
    }

  }

}
