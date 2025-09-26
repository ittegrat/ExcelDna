using System;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;

namespace ExcelDnaUnpack
{

  internal static partial class NativeMethods
  {

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool FreeLibrary(IntPtr hModule);

  }

  internal sealed class SafeModuleHandle : SafeHandle
  {

    public SafeModuleHandle()
      : base(IntPtr.Zero, true) { }

    public override bool IsInvalid {
#if NETFRAMEWORK
      [ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)]
      [PrePrepareMethod]
#endif
      get { return (handle == IntPtr.Zero); }
    }

#if NETFRAMEWORK
    [ReliabilityContract(Consistency.WillNotCorruptState, Cer.MayFail)]
    [PrePrepareMethod]
#endif
    protected override bool ReleaseHandle() {
      return NativeMethods.FreeLibrary(handle);
    }

  }

}
