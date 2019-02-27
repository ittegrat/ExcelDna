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
      [ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)]
      [PrePrepareMethod]
      get { return (handle == IntPtr.Zero); }
    }

    [ReliabilityContract(Consistency.WillNotCorruptState, Cer.MayFail)]
    [PrePrepareMethod]
    protected override bool ReleaseHandle() {
      return NativeMethods.FreeLibrary(handle);
    }

  }

}
