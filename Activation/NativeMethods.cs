using System;
using System.IO;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace Activation
{
    internal static class NativeMethods
    {
        [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern IntPtr LoadLibrary(string lpFileName);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool FreeLibrary(IntPtr hModule);

        #region Microsoft Software Protection Service
        internal enum SLDATATYPE
        {
            // ReSharper disable InconsistentNaming
            SL_DATA_NONE = 0,
            SL_DATA_SZ = 1,
            SL_DATA_DWORD = 4,
            SL_DATA_BINARY = 3,
            SL_DATA_MULTI_SZ = 7,
            SL_DATA_SUM = 100,
            // ReSharper restore InconsistentNaming
        }

        #region Microsoft Office Software Protection Service
        [DllImport("osppc.dll", EntryPoint = "SLOpen", CharSet = CharSet.Unicode)]
        internal static extern int OSPPOpen(ref IntPtr phSLC);

        [DllImport("osppc.dll", EntryPoint = "SLClose", CharSet = CharSet.Unicode)]
        internal static extern int OSPPClose(IntPtr hSLC);

        [DllImport("osppc.dll", EntryPoint = "SLGetApplicationInformation", CharSet = CharSet.Unicode)]
        internal static extern int OSPPGetApplicationInformation(IntPtr hSLC, ref Guid pApplicationId, [InAttribute] [MarshalAsAttribute(UnmanagedType.LPWStr)] string pwszValueName, ref SLDATATYPE peDataType, ref uint pcbValue, ref byte[] ppbValue);
        #endregion

        #region Microsoft Windows Software Protection Service
        [DllImport("Slc.dll", EntryPoint = "SLOpen", CharSet = CharSet.Unicode)]
        internal static extern int SPPOpen(ref IntPtr phSLC);

        [DllImport("Slc.dll", EntryPoint = "SLClose", CharSet = CharSet.Unicode)]
        internal static extern int SPPClose(IntPtr hSLC);

        [DllImport("Slc.dll", EntryPoint = "SLGetApplicationInformation", CharSet = CharSet.Unicode)]
        internal static extern int SPPGetApplicationInformation(IntPtr hSLC, ref Guid pApplicationId, [InAttribute] [MarshalAsAttribute(UnmanagedType.LPWStr)] string pwszValueName, ref SLDATATYPE peDataType, ref uint pcbValue, ref byte[] ppbValue);
        #endregion
        #endregion

        #region TAP Adapter Control Functions
        [DllImport("Kernel32.dll", /* ExactSpelling = true, */ SetLastError = true, CharSet = CharSet.Auto)]
        internal static extern SafeFileHandle CreateFile(string filename, [MarshalAs(UnmanagedType.U4)] FileAccess fileaccess, [MarshalAs(UnmanagedType.U4)] FileShare fileshare, int securityattributes, [MarshalAs(UnmanagedType.U4)] FileMode creationdisposition, int flags, IntPtr template);

        [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true, CharSet = CharSet.Auto)]
        internal static extern bool DeviceIoControl(SafeFileHandle hDevice, uint dwIoControlCode, IntPtr lpInBuffer, uint nInBufferSize, IntPtr lpOutBuffer, uint nOutBufferSize, out int lpBytesReturned, IntPtr lpOverlapped);
        #endregion

        #region Device Manager Functions
        [DllImport("setupapi.dll", CallingConvention = CallingConvention.Winapi, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool SetupDiCallClassInstaller(DiFunction installFunction, SafeDeviceInfoSetHandle deviceInfoSet, [In] ref DeviceInfoData deviceInfoData);

        [DllImport("setupapi.dll", CallingConvention = CallingConvention.Winapi, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool SetupDiEnumDeviceInfo(SafeDeviceInfoSetHandle deviceInfoSet, int memberIndex, ref DeviceInfoData deviceInfoData);

        [DllImport("setupapi.dll", CallingConvention = CallingConvention.Winapi, CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern SafeDeviceInfoSetHandle SetupDiGetClassDevs([In] ref Guid classGuid, [MarshalAs(UnmanagedType.LPWStr)] string enumerator, IntPtr hwndParent, SetupDiGetClassDevsFlags flags);

        [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Auto)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool SetupDiGetDeviceInstanceId(IntPtr deviceInfoSet, ref DeviceInfoData did, [MarshalAs(UnmanagedType.LPTStr)] StringBuilder deviceInstanceId, int deviceInstanceIdSize, out int requiredSize);

        [SuppressUnmanagedCodeSecurity]
        [ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)]
        [DllImport("setupapi.dll", CallingConvention = CallingConvention.Winapi, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool SetupDiDestroyDeviceInfoList(IntPtr deviceInfoSet);

        [DllImport("setupapi.dll", CallingConvention = CallingConvention.Winapi, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool SetupDiSetClassInstallParams(SafeDeviceInfoSetHandle deviceInfoSet, [In] ref DeviceInfoData deviceInfoData, [In] ref PropertyChangeParameters classInstallParams, int classInstallParamsSize);

        [DllImport("setupapi.dll", CallingConvention = CallingConvention.Winapi, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool SetupDiRestartDevices(SafeDeviceInfoSetHandle deviceInfoSet, [In] ref DeviceInfoData deviceInfoData);


        /* TODO: INF Install
        [DllImport("DIFxAPI.dll", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern Int32 DriverPackageInstall([MarshalAs(UnmanagedType.LPTStr)] string driverPackageInfPath, Int32 flags, IntPtr pInstallerInfo, out bool pNeedReboot);

        [DllImport("setupapi.dll")]
        internal static extern bool SetupCopyOEMInf(string sourceInfFileName, string oemSourceMediaLocation, int oemSourceMediaType, int copyStyle, string destinationInfFileName, int destinationInfFileNameSize, int requiredSize, string destinationInfFileNameComponent);

        [DllImport("newdev.dll")]
        internal static extern bool UpdateDriverForPlugAndPlayDevices(IntPtr hwndParent, string hardwareId, string fullInfPath, uint installFlags, bool bRebootRequired);
         */
        #endregion

        #region IP Helper Route Control Functions
        [StructLayoutAttribute(LayoutKind.Sequential)]
        public struct MibIPForwardRow
        {
            public uint dwForwardDest;
            public uint dwForwardMask;
            public uint dwForwardPolicy;
            public uint dwForwardNextHop;
            public uint dwForwardIfIndex;
            public uint dwForwardType;
            public uint dwForwardProto;
            public uint dwForwardAge;
            public uint dwForwardNextHopAS;
            public uint dwForwardMetric1;
            public uint dwForwardMetric2;
            public uint dwForwardMetric3;
            public uint dwForwardMetric4;
            public uint dwForwardMetric5;
        }

        public enum NlRouteProtocol
        {
            RouteProtocolOther = 1,
            RouteProtocolLocal,
            RouteProtocolNetMgmt,
            RouteProtocolIcmp,
            RouteProtocolEgp,
            RouteProtocolGgp,
            RouteProtocolHello,
            RouteProtocolRip,
            RouteProtocolIsIs,
            RouteProtocolEsIs,
            RouteProtocolCisco,
            RouteProtocolBbn,
            RouteProtocolOspf,
            RouteProtocolBgp,
            MibIPProtoOther = 1,
            MibIPProtoLocal = 2,
            MibIPProtoNetmgmt = 3,
            MibIPProtoICMP = 4,
            MibIProtoEGP = 5,
            MibIPProtoGGP = 6,
            MibIPProtoHello = 7,
            MibIPProtoRIP = 8,
            MibIPProtoIsIs = 9,
            MibIPProtoEsIs = 10,
            MibIPProtoCisco = 11,
            MibIPProtoBBN = 12,
            MibIPProtoOSPF = 13,
            MibIPProtoBGP = 14,
            MibIPProtoNTAutostatic = 10002,
            MibIPProtoNTStatic = 10006,
            MibIPProtoNTStaticNonDOD = 10007,
        }

        [DllImportAttribute("Iphlpapi.dll", EntryPoint = "CreateIpForwardEntry")]
        internal static extern uint CreateIpForwardEntry(ref MibIPForwardRow pRoute);

        [DllImportAttribute("Iphlpapi.dll", EntryPoint = "DeleteIpForwardEntry")]
        internal static extern uint DeleteIpForwardEntry(ref MibIPForwardRow pRoute);
        #endregion

        #region DLL Injection Functions
        [Flags]
        public enum AllocationType
        {
            Commit = 0x1000,
            Reserve = 0x2000,
            Decommit = 0x4000,
            Release = 0x8000,
            Reset = 0x80000,
            Physical = 0x400000,
            TopDown = 0x100000,
            WriteWatch = 0x200000,
            LargePages = 0x20000000
        }

        [Flags]
        public enum MemoryProtection
        {
            Execute = 0x10,
            ExecuteRead = 0x20,
            ExecuteReadWrite = 0x40,
            ExecuteWriteCopy = 0x80,
            NoAccess = 0x01,
            ReadOnly = 0x02,
            ReadWrite = 0x04,
            WriteCopy = 0x08,
            GuardModifierflag = 0x100,
            NoCacheModifierflag = 0x200,
            WriteCombineModifierflag = 0x400
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr OpenProcess(uint dwDesiredAccess, int bInheritHandle, uint dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern int CloseHandle(IntPtr hObject);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr GetProcAddress(IntPtr hModule, string lpProcName);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("kernel32.dll", SetLastError = true, ExactSpelling = true)]
        internal static extern IntPtr VirtualAllocEx(IntPtr hProcess, IntPtr lpAddress, uint dwSize, AllocationType flAllocationType, MemoryProtection flProtect);

        [DllImport("kernel32.dll", SetLastError = true, ExactSpelling = true)]
        internal static extern bool VirtualFreeEx(IntPtr hProcess, IntPtr lpAddress, int dwSize, AllocationType dwFreeType);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern int WriteProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, byte[] buffer, uint size, int lpNumberOfBytesWritten);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern bool WriteProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, IntPtr lpBuffer, uint nSize, out int lpNumberOfBytesWritten);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr CreateRemoteThread(IntPtr hProcess, IntPtr lpThreadAttribute, uint dwStackSize, IntPtr lpStartAddress, IntPtr lpParameter, uint dwCreationFlags, IntPtr lpThreadId);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern int WaitForSingleObject(IntPtr handle, uint wait);
        #endregion
    }
}