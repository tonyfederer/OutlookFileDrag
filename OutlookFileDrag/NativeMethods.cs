using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace OutlookFileDrag
{
    static class NativeMethods
    {
        public const int S_OK = 0;
        public const int S_FALSE = 1;

        public const short CF_TEXT = 1;
        public const short CF_UNICODETEXT = 13;
        public const short CF_HDROP = 15;

        public const int DATA_S_SAMEFORMATETC = 0x00040130;

        public const int DRAGDROP_S_DROP = 0x00040100;
        public const int DRAGDROP_S_CANCEL = 0x00040101;
        public const int DRAGDROP_S_USEDEFAULTCURSORS = 0x00040102;

        public const uint DROPEFFECT_NONE = 0;
        public const uint DROPEFFECT_COPY = 1;
        public const uint DROPEFFECT_MOVE = 2;
        public const uint DROPEFFECT_LINK = 4;
        public const uint DROPEFFECT_SCROLL = 0x80000000;

        public const int DV_E_LINDEX = unchecked((int)0x80040068);
        public const int DV_E_FORMATETC = unchecked((int)0x80040064);
        public const int DV_E_TYMED = unchecked((int)0x80040069);
        public const int DV_E_DVASPECT = unchecked((int)0x8004006B);

        public const int E_INVALIDARG = unchecked((int)0x80070057);
        public const int E_FAIL = unchecked((int)0x80004005);
        public const int E_NOTIMPL = unchecked((int)0x80004001);
        public const int E_OUTOFMEMORY = unchecked((int)0x8007000E);
        public const int E_UNSPEC = E_FAIL;
        public const int E_UNEXPECTED = unchecked((int)0x8000FFFF);

        public const int MAX_PATH = 260;

        [DllImport("ole32.dll")]
        public static extern int DoDragDrop(NativeMethods.IDataObject pDataObj, IntPtr pDropSource, uint dwOKEffects, out uint pdwEffect);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr GlobalLock(IntPtr handle);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public static extern bool GlobalUnlock(IntPtr handle);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public static extern int GlobalSize(IntPtr handle);

        [DllImport("ole32.dll", PreserveSig = false)]
        public static extern ILockBytes CreateILockBytesOnHGlobal(IntPtr hGlobal, bool fDeleteOnRelease);

        [DllImport("ole32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        public static extern IStorage StgCreateDocfileOnILockBytes(ILockBytes plkbyt, uint grfMode, uint reserved);

        [DllImport("ole32.dll")]
        internal static extern void ReleaseStgMedium(ref STGMEDIUM medium);

        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000000B-0000-0000-C000-000000000046")]
        public interface IStorage
        {
            [return: MarshalAs(UnmanagedType.Interface)]
            IStream CreateStream([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In, MarshalAs(UnmanagedType.U4)] int grfMode, [In, MarshalAs(UnmanagedType.U4)] int reserved1, [In, MarshalAs(UnmanagedType.U4)] int reserved2);
            [return: MarshalAs(UnmanagedType.Interface)]
            IStream OpenStream([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, IntPtr reserved1, [In, MarshalAs(UnmanagedType.U4)] int grfMode, [In, MarshalAs(UnmanagedType.U4)] int reserved2);
            [return: MarshalAs(UnmanagedType.Interface)]
            IStorage CreateStorage([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In, MarshalAs(UnmanagedType.U4)] int grfMode, [In, MarshalAs(UnmanagedType.U4)] int reserved1, [In, MarshalAs(UnmanagedType.U4)] int reserved2);
            [return: MarshalAs(UnmanagedType.Interface)]
            IStorage OpenStorage([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, IntPtr pstgPriority, [In, MarshalAs(UnmanagedType.U4)] int grfMode, IntPtr snbExclude, [In, MarshalAs(UnmanagedType.U4)] int reserved);
            void CopyTo(int ciidExclude, [In, MarshalAs(UnmanagedType.LPArray)] Guid[] pIIDExclude, IntPtr snbExclude, [In, MarshalAs(UnmanagedType.Interface)] IStorage stgDest);
            void MoveElementTo([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In, MarshalAs(UnmanagedType.Interface)] IStorage stgDest, [In, MarshalAs(UnmanagedType.BStr)] string pwcsNewName, [In, MarshalAs(UnmanagedType.U4)] int grfFlags);
            void Commit(int grfCommitFlags);
            void Revert();
            void EnumElements([In, MarshalAs(UnmanagedType.U4)] int reserved1, IntPtr reserved2, [In, MarshalAs(UnmanagedType.U4)] int reserved3, [MarshalAs(UnmanagedType.Interface)] out object ppVal);
            void DestroyElement([In, MarshalAs(UnmanagedType.BStr)] string pwcsName);
            void RenameElement([In, MarshalAs(UnmanagedType.BStr)] string pwcsOldName, [In, MarshalAs(UnmanagedType.BStr)] string pwcsNewName);
            void SetElementTimes([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In] System.Runtime.InteropServices.ComTypes.FILETIME pctime, [In] System.Runtime.InteropServices.ComTypes.FILETIME patime, [In] System.Runtime.InteropServices.ComTypes.FILETIME pmtime);
            void SetClass([In] ref Guid clsid);
            void SetStateBits(int grfStateBits, int grfMask);
            void Stat([Out]out System.Runtime.InteropServices.ComTypes.STATSTG pStatStg, int grfStatFlag);
        }

        [ComImport, Guid("0000000A-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface ILockBytes
        {
            void ReadAt([In, MarshalAs(UnmanagedType.U8)] long ulOffset, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pv, [In, MarshalAs(UnmanagedType.U4)] int cb, IntPtr pcbRead);
            void WriteAt([In, MarshalAs(UnmanagedType.U8)] long ulOffset, IntPtr pv, [In, MarshalAs(UnmanagedType.U4)] int cb, [Out, MarshalAs(UnmanagedType.LPArray)] int[] pcbWritten);
            void Flush();
            void SetSize([In, MarshalAs(UnmanagedType.U8)] long cb);
            void LockRegion([In, MarshalAs(UnmanagedType.U8)] long libOffset, [In, MarshalAs(UnmanagedType.U8)] long cb, [In, MarshalAs(UnmanagedType.U4)] int dwLockType);
            void UnlockRegion([In, MarshalAs(UnmanagedType.U8)] long libOffset, [In, MarshalAs(UnmanagedType.U8)] long cb, [In, MarshalAs(UnmanagedType.U4)] int dwLockType);
            void Stat([Out]out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg, [In, MarshalAs(UnmanagedType.U4)] int grfStatFlag);
        }

        [ComImport, Guid("0000010E-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IDataObject
        {
            [PreserveSig]
            int GetData([In] ref FORMATETC format, out STGMEDIUM medium);
            [PreserveSig]
            int GetDataHere([In] ref FORMATETC format, ref STGMEDIUM medium);
            [PreserveSig]
            int QueryGetData([In] ref FORMATETC format);
            [PreserveSig]
            int GetCanonicalFormatEtc([In] ref FORMATETC formatIn, out FORMATETC formatOut);
            [PreserveSig]
            int SetData([In] ref FORMATETC formatIn, [In] ref STGMEDIUM medium, [MarshalAs(UnmanagedType.Bool)] bool release);
            [PreserveSig]
            int EnumFormatEtc(DATADIR direction, out System.Runtime.InteropServices.ComTypes.IEnumFORMATETC ppenumFormatEtc);
            [PreserveSig]
            int DAdvise([In] ref FORMATETC pFormatetc, ADVF advf, IAdviseSink adviseSink, out int connection);
            [PreserveSig]
            int DUnadvise(int connection);
            [PreserveSig]
            int EnumDAdvise(out IEnumSTATDATA enumAdvise);
        }

        [ComImport, Guid("00000121-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IDropSource
        {
            [PreserveSig]
            int QueryContinueDrag([MarshalAs(UnmanagedType.Bool)] bool fEscapePressed, int grfKeyState);
            [PreserveSig]
            int GiveFeedback(int dwEffect);
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public sealed class FILEGROUPDESCRIPTORA
        {
            public uint cItems;
            //public FILEDESCRIPTORA[] fgd;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        public sealed class FILEDESCRIPTORA
        {
            public uint dwFlags;
            public Guid clsid;
            public SIZEL sizel;
            public POINTL pointl;
            public uint dwFileAttributes;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
            public uint nFileSizeHigh;
            public uint nFileSizeLow;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string cFileName;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public sealed class FILEGROUPDESCRIPTORW
        {
            public uint cItems;
            //public FILEDESCRIPTORW[] fgd;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public sealed class FILEDESCRIPTORW
        {
            public uint dwFlags;
            public Guid clsid;
            public SIZEL sizel;
            public POINTL pointl;
            public uint dwFileAttributes;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
            public uint nFileSizeHigh;
            public uint nFileSizeLow;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string cFileName;
        }

        [StructLayout(LayoutKind.Sequential)]
        public sealed class POINTL
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public sealed class SIZEL
        {
            public int cx;
            public int cy;
        }

        [StructLayout(LayoutKind.Sequential)]
        public sealed class DROPFILES
        {
            public int pFiles;
            public int X;
            public int Y;
            public bool fNC;
            public bool fWide;
        }
    
        [UnmanagedFunctionPointer(CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.I4)]
        public delegate int DragDropDelegate(NativeMethods.IDataObject pDataObj, IntPtr pDropSource, uint dwOKEffects, out uint pdwEffect);

    }
}
