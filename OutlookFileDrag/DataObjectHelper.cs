using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace OutlookFileDrag
{
    static class DataObjectHelper
    {
        private static readonly byte[] serializedObjectID = new Guid("FD9EA796-3B13-4370-A679-56106BB288FB").ToByteArray();

        internal static int GetClipboardFormat(string name)
        {
            var format = System.Windows.Forms.DataFormats.GetFormat(name);
            if (format == null)
                return 0;
            else
                return format.Id;
        }

        internal static bool GetDataPresent(NativeMethods.IDataObject data, string formatName)
        {
            //Check if drag contains virtual files
            FORMATETC format = new FORMATETC();
            format.cfFormat = (short)GetClipboardFormat("FileGroupDescriptorW");
            format.dwAspect = DVASPECT.DVASPECT_CONTENT;
            format.lindex = -1;
            format.ptd = IntPtr.Zero;
            format.tymed = TYMED.TYMED_ISTREAM | TYMED.TYMED_ISTORAGE | TYMED.TYMED_HGLOBAL;

            return data.QueryGetData(format) == NativeMethods.S_OK;
        }

        internal static void SetDropFiles(ref STGMEDIUM medium, string[] filenames)
        {
            //Create DROPFILES structure
            NativeMethods.DROPFILES dropFiles = new NativeMethods.DROPFILES();
            dropFiles.pFiles = Marshal.SizeOf(dropFiles);
            dropFiles.fWide = true;     //Unicode

            //Get null-separated list of filenames terminated with double null
            string filenameList = string.Join("\0", filenames) + "\0\0";
            byte[] filenameBytes = System.Text.Encoding.Unicode.GetBytes(filenameList);

            //Allocate global memory and get pointer
            int dataLength = Marshal.SizeOf(dropFiles) + filenameBytes.Length;
            IntPtr ptrDropFiles = Marshal.AllocHGlobal(dataLength);
            if (ptrDropFiles == IntPtr.Zero)
                throw new OutOfMemoryException();

            //Copy DROPFILES structure to global memory.
            Marshal.StructureToPtr(dropFiles, ptrDropFiles, true);

            //Copy filenames to memory after DROPFILES structure
            IntPtr ptrFiles = IntPtr.Add(ptrDropFiles, Marshal.SizeOf(dropFiles));
            Marshal.Copy(filenameBytes, 0, ptrFiles, filenameBytes.Length);
            
            //Load structure into medium
            medium.unionmember = ptrDropFiles;
            medium.tymed = TYMED.TYMED_HGLOBAL;
        }

        internal static string[] GetFilenames(NativeMethods.IDataObject data)
        {
            //Try Unicode first
            string[] filenames = GetFilenamesUnicode(data);

            //If Unicode returns null, try ANSI
            if (filenames == null)
                filenames = GetFilenamesAnsi(data);

            return filenames;
        }

        internal static string[] GetFilenamesAnsi(NativeMethods.IDataObject data)
        {
            IntPtr fgdaPtr = IntPtr.Zero;

            try
            {
                //Define FileGroupDescriptor format
                FORMATETC format = new FORMATETC();
                format.cfFormat = (short)System.Windows.Forms.DataFormats.GetFormat("FileGroupDescriptor").Id;
                format.dwAspect = DVASPECT.DVASPECT_CONTENT;
                format.lindex = -1;
                format.ptd = IntPtr.Zero;
                format.tymed = TYMED.TYMED_ISTREAM | TYMED.TYMED_ISTORAGE | TYMED.TYMED_HGLOBAL;

                //Query if format exists in data
                if (data.QueryGetData(format) != NativeMethods.S_OK)
                    return null;

                //Get data into medium
                STGMEDIUM medium = new STGMEDIUM();
                data.GetData(format, out medium);

                //Read medium into byte array
                byte[] bytes;
                using (MemoryStream stream = new MemoryStream())
                {
                    DataObjectHelper.ReadMediumIntoStream(medium, stream);
                    bytes = new byte[stream.Length];
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.Read(bytes, 0, bytes.Length);
                }

                //Copy the file group descriptor into unmanaged memory
                fgdaPtr = Marshal.AllocHGlobal(bytes.Length);
                Marshal.Copy(bytes, 0, fgdaPtr, bytes.Length);

                //Marshal the unmanaged memory to a FILEGROUPDESCRIPTORA struct
                object fgdObj = Marshal.PtrToStructure(fgdaPtr, typeof(NativeMethods.FILEGROUPDESCRIPTORA));
                NativeMethods.FILEGROUPDESCRIPTORA fgd = (NativeMethods.FILEGROUPDESCRIPTORA)fgdObj;

                //Create an array to store file names
                string[] filenames = new string[fgd.cItems];

                //Get the pointer to the first file descriptor
                IntPtr fdPtr = IntPtr.Add(fgdaPtr, sizeof(uint));

                //Loop for the number of files acording to the file group descriptor
                for (int fdIndex = 0; fdIndex < fgd.cItems; fdIndex++)
                {
                    //Marshal the pointer to the file descriptor as a FILEDESCRIPTORA struct
                    object fdObj = Marshal.PtrToStructure(fdPtr, typeof(NativeMethods.FILEDESCRIPTORA));
                    NativeMethods.FILEDESCRIPTORA fd = (NativeMethods.FILEDESCRIPTORA)fdObj;

                    //Get filename of file descriptor and put in array
                    filenames[fdIndex] = fd.cFileName;

                    //Move the file descriptor pointer to the next file descriptor
                    fdPtr = IntPtr.Add(fdPtr, Marshal.SizeOf(fd));
                }

                return filenames;

            }
            finally
            {
                Marshal.FreeHGlobal(fgdaPtr);		
            }
        }

        internal static string[] GetFilenamesUnicode(NativeMethods.IDataObject data)
        {
            IntPtr fgdaPtr = IntPtr.Zero;

            try
            {
                //Define FileGroupDescriptorW format
                FORMATETC format = new FORMATETC();
                format.cfFormat = (short)System.Windows.Forms.DataFormats.GetFormat("FileGroupDescriptorW").Id;
                format.dwAspect = DVASPECT.DVASPECT_CONTENT;
                format.lindex = -1;
                format.ptd = IntPtr.Zero;
                format.tymed = TYMED.TYMED_ISTREAM | TYMED.TYMED_ISTORAGE | TYMED.TYMED_HGLOBAL;

                //Query if format exists in data
                if (data.QueryGetData(format) != NativeMethods.S_OK)
                    return null;

                //Get data into medium
                STGMEDIUM medium = new STGMEDIUM();
                data.GetData(format, out medium);

                //Read medium into string
                byte[] bytes;
                using (MemoryStream stream = new MemoryStream())
                {
                    DataObjectHelper.ReadMediumIntoStream(medium, stream);
                    bytes = new byte[stream.Length];
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.Read(bytes, 0, bytes.Length);
                }

                //Copy the file group descriptor into unmanaged memory
                fgdaPtr = Marshal.AllocHGlobal(bytes.Length);
                if (fgdaPtr == IntPtr.Zero)
                    throw new OutOfMemoryException();
                Marshal.Copy(bytes, 0, fgdaPtr, bytes.Length);

                //Marshal the unmanaged memory to a FILEGROUPDESCRIPTORW struct
                object fgdObj = Marshal.PtrToStructure(fgdaPtr, typeof(NativeMethods.FILEGROUPDESCRIPTORW));
                NativeMethods.FILEGROUPDESCRIPTORW fgd = (NativeMethods.FILEGROUPDESCRIPTORW)fgdObj;

                //Create an array to store file names
                string[] filenames = new string[fgd.cItems];

                //Get the pointer to the first file descriptor
                IntPtr fdPtr = IntPtr.Add(fgdaPtr, sizeof(uint));

                //Loop for the number of files acording to the file group descriptor
                for (int fdIndex = 0; fdIndex < fgd.cItems; fdIndex++)
                {
                    //Marshal the pointer to the file descriptor as a FILEDESCRIPTORW struct
                    object fdObj = Marshal.PtrToStructure(fdPtr, typeof(NativeMethods.FILEDESCRIPTORW));
                    NativeMethods.FILEDESCRIPTORW fd = (NativeMethods.FILEDESCRIPTORW)fdObj;

                    //Get filename of file descriptor and put in array
                    filenames[fdIndex] = fd.cFileName;

                    //Move the file descriptor pointer to the next file descriptor
                    fdPtr = IntPtr.Add(fdPtr, Marshal.SizeOf(fd));
                }

                return filenames;

            }
            finally
            {
                Marshal.FreeHGlobal(fgdaPtr);
            }
        }

        internal static void ReadFileContents(NativeMethods.IDataObject data, int index, Stream stream)
        {
            //Define FileContents format
            FORMATETC format = new FORMATETC();
            format.cfFormat = (short)System.Windows.Forms.DataFormats.GetFormat("FileContents").Id;
            format.dwAspect = DVASPECT.DVASPECT_CONTENT;
            format.lindex = index;
            format.ptd = IntPtr.Zero;
            format.tymed = TYMED.TYMED_ISTREAM | TYMED.TYMED_ISTORAGE | TYMED.TYMED_HGLOBAL;

            //Get data
            STGMEDIUM medium = new STGMEDIUM();
            data.GetData(format, out medium);
            
            //Read medium into stream
            ReadMediumIntoStream(medium, stream);
        }

        internal static void ReadMediumIntoStream(STGMEDIUM medium, Stream stream)
        {
            switch (medium.tymed)
            {
                case TYMED.TYMED_ISTREAM:
                    ReadIStreamIntoStream(medium.unionmember, stream);
                    break;
                case TYMED.TYMED_ISTORAGE:
                    ReadIStorageIntoStream(medium.unionmember, stream);
                    break;
                case TYMED.TYMED_HGLOBAL:
                    ReadHGlobalIntoStream(medium.unionmember, stream);
                    break;
                default:
                    throw new NotImplementedException(string.Format("Cannot read medium type {0}", medium.tymed));
            }
        }

        private static void ReadIStorageIntoStream(IntPtr handle, Stream stream)
        {
            //To handle a IStorage it needs to be written into a second unmanaged memory mapped storage 
            //and then the data can be read from memory into a managed byte and returned as a MemoryStream

            NativeMethods.IStorage iStorage = null;
            NativeMethods.IStorage iStorage2 = null;
            NativeMethods.ILockBytes iLockBytes = null;
            System.Runtime.InteropServices.ComTypes.STATSTG iLockBytesStat;
            try
            {
                //Marshal the returned pointer to a IStorage object
                iStorage = (NativeMethods.IStorage)Marshal.GetObjectForIUnknown(handle);
                Marshal.Release(handle);

                //Create a ILockBytes (unmanaged byte array) and then create a IStorage using the byte array as a backing store
                iLockBytes = NativeMethods.CreateILockBytesOnHGlobal(IntPtr.Zero, true);
                iStorage2 = NativeMethods.StgCreateDocfileOnILockBytes(iLockBytes, 0x00001012, 0);

                //Copy the returned IStorage into the new IStorage
                iStorage.CopyTo(0, null, IntPtr.Zero, iStorage2);
                iLockBytes.Flush();
                iStorage2.Commit(0);

                //Get the STATSTG of the ILockBytes to determine how many bytes were written to it
                iLockBytesStat = new System.Runtime.InteropServices.ComTypes.STATSTG();
                iLockBytes.Stat(out iLockBytesStat, 1);
                int iLockBytesSize = (int)iLockBytesStat.cbSize;

                //Read the data from the ILockBytes (unmanaged byte array) into a managed byte array
                //byte[] iLockBytesContent = new byte[iLockBytesSize];
                //iLockBytes.ReadAt(0, iLockBytesContent, iLockBytesContent.Length, null);

                //Read bytes into stream
                IntPtr ptrRead = Marshal.AllocCoTaskMem(sizeof(int));
                byte[] buffer = new byte[1024];
                int offset = 0;
                int bytesRead;
                while (true)
                {
                    iLockBytes.ReadAt(offset, buffer, buffer.Length, ptrRead);
                    bytesRead = Marshal.ReadInt32(ptrRead);
                    if (bytesRead == 0)
                        break;
                    stream.Write(buffer, 0, bytesRead);
                    offset += bytesRead;
                }
                stream.Seek(0, SeekOrigin.Begin);

                //Wrap the managed byte array into a memory stream and return it
                //return new MemoryStream(iLockBytesContent);
            }
            finally
            {
                //release all unmanaged objects
                Marshal.ReleaseComObject(iStorage2);
                Marshal.ReleaseComObject(iLockBytes);
                Marshal.ReleaseComObject(iStorage);
            }
        }

        private static void ReadIStreamIntoStream(IntPtr handle, Stream stream)
        {
            IStream iStream = null;
            System.Runtime.InteropServices.ComTypes.STATSTG iStreamStat;
            try
            {
                //Marshal the returned pointer to a IStream object
                iStream = (IStream)Marshal.GetObjectForIUnknown(handle);
                Marshal.Release(handle);

                //Get the STATSTG of the IStream to determine how many bytes are in it
                iStreamStat = new System.Runtime.InteropServices.ComTypes.STATSTG();
                iStream.Stat(out iStreamStat, 0);
                int iStreamSize = (int)iStreamStat.cbSize;

                IntPtr ptrRead = Marshal.AllocCoTaskMem(sizeof(int));
                byte[] buffer = new byte[1024];
                int bytesRead;
                while (true)
                {
                    iStream.Read(buffer, buffer.Length, ptrRead);
                    bytesRead = Marshal.ReadInt32(ptrRead);
                    if (bytesRead == 0)
                        break;
                    else
                        stream.Write(buffer, 0, bytesRead);
                }
                stream.Seek(0, SeekOrigin.Begin);
            }
            finally
            {
                //Release all unmanaged objects
                Marshal.ReleaseComObject(iStream);
            }
        }

        private static void ReadHGlobalIntoStream(IntPtr handle, Stream stream)
        {
            IntPtr source = NativeMethods.GlobalLock(new HandleRef((object)null, handle));
            if (source == IntPtr.Zero)
                throw new ExternalException("An external error occurred in GlobalLock", -2147024882);
            try
            {
                int length = NativeMethods.GlobalSize(new HandleRef((object)null, handle));
                byte[] buffer = new byte[length];
                Marshal.Copy(source, buffer, 0, length);
                stream.Write(buffer, 0, buffer.Length);
            }
            finally
            {
                //Release all unmanaged objects
                NativeMethods.GlobalUnlock(new HandleRef((object)null, handle));
            }
        }
    }
}
