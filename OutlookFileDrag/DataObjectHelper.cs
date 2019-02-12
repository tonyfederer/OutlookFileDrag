using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using log4net;

namespace OutlookFileDrag
{
    static class DataObjectHelper
    {
        private static ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

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
            //Check if data is present
            log.DebugFormat("Get data present: {0}", formatName);
            FORMATETC format = new FORMATETC();
            format.cfFormat = (short)GetClipboardFormat(formatName);
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

            //Copy DROPFILES structure to global memory.
            Marshal.StructureToPtr(dropFiles, ptrDropFiles, true);

            //Copy filenames to memory after DROPFILES structure
            IntPtr ptrFiles = IntPtr.Add(ptrDropFiles, Marshal.SizeOf(dropFiles));
            Marshal.Copy(filenameBytes, 0, ptrFiles, filenameBytes.Length);
            
            //Load structure into medium
            medium.unionmember = ptrDropFiles;
            medium.tymed = TYMED.TYMED_HGLOBAL;
            medium.pUnkForRelease = IntPtr.Zero;        //HGLOBAL to be released by caller
        }

        internal static string[] GetFilenames(NativeMethods.IDataObject data)
        {
            log.Debug("Getfilenames");

            //Try Unicode first
            string[] filenames = GetFilenamesUnicode(data);

            //If Unicode returns null, try ANSI
            if (filenames == null)
                filenames = GetFilenamesAnsi(data);

            return filenames;
        }

        internal static string[] GetFilenamesAnsi(NativeMethods.IDataObject data)
        {
            log.Debug("Getting filenames (ANSI)");
            IntPtr ptrFgd = IntPtr.Zero;
            STGMEDIUM medium = new STGMEDIUM();

            try
            {
                //Define FileGroupDescriptor format
                FORMATETC format = new FORMATETC();
                format.cfFormat = (short)GetClipboardFormat("FileGroupDescriptor");
                format.dwAspect = DVASPECT.DVASPECT_CONTENT;
                format.lindex = -1;
                format.ptd = IntPtr.Zero;
                format.tymed = TYMED.TYMED_ISTREAM | TYMED.TYMED_ISTORAGE | TYMED.TYMED_HGLOBAL;

                //Query if format exists in data
                if (data.QueryGetData(format) != NativeMethods.S_OK)
                {
                    log.Debug("No filenames found");
                    return null;
                }

                //Get data into medium
                int retVal = data.GetData(format, out medium);
                if (retVal != NativeMethods.S_OK)
                    throw new Exception(string.Format("Could not get FileGroupDescriptor format.  Error returned: {0}", retVal));

                //Read medium into byte array
                log.Debug("Reading structure into memory stream");
                byte[] bytes;
                using (MemoryStream stream = new MemoryStream())
                {
                    DataObjectHelper.ReadMediumIntoStream(medium, stream);
                    bytes = new byte[stream.Length];
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.Read(bytes, 0, bytes.Length);
                }

                //Copy byte array into unmanaged memory
                log.Debug("Copying structure into unmanaged memory");
                ptrFgd = Marshal.AllocHGlobal(bytes.Length);
                Marshal.Copy(bytes, 0, ptrFgd, bytes.Length);

                //Marshal unmanaged memory to a FILEGROUPDESCRIPTORA struct
                log.Debug("Marshaling unmanaged memory into FILEGROUPDESCRIPTORA struct");
                NativeMethods.FILEGROUPDESCRIPTORA fgd = (NativeMethods.FILEGROUPDESCRIPTORA)Marshal.PtrToStructure(ptrFgd, typeof(NativeMethods.FILEGROUPDESCRIPTORA));
                log.Debug(string.Format("Files found: {0}", fgd.cItems));

                //Create an array to store file names
                string[] filenames = new string[fgd.cItems];

                //Get the pointer to the first file descriptor
                IntPtr fdPtr = IntPtr.Add(ptrFgd, sizeof(uint));

                //Loop for the number of files acording to the file group descriptor
                for (int fdIndex = 0; fdIndex < fgd.cItems; fdIndex++)
                {
                    log.DebugFormat("Filenames found: {0}", string.Join(", ", filenames));

                    //Marshal pointer to a FILEDESCRIPTORA struct
                    NativeMethods.FILEDESCRIPTORA fd = (NativeMethods.FILEDESCRIPTORA)Marshal.PtrToStructure(fdPtr, typeof(NativeMethods.FILEDESCRIPTORA));

                    //Get filename of file descriptor and put in array
                    filenames[fdIndex] = fd.cFileName;

                    //Move the file descriptor pointer to the next file descriptor
                    fdPtr = IntPtr.Add(fdPtr, Marshal.SizeOf(fd));
                }

                log.DebugFormat("Filenames found: {0}", string.Join(", ", filenames));

                return filenames;

            }
            finally
            {
                //Release all unmanaged objects
                Marshal.FreeHGlobal(ptrFgd);
                if (medium.pUnkForRelease == null)
                    NativeMethods.ReleaseStgMedium(ref medium);
            }
        }

        internal static string[] GetFilenamesUnicode(NativeMethods.IDataObject data)
        {
            log.Debug("Getting filenames (Unicode)");
            IntPtr ptrFgd = IntPtr.Zero;
            STGMEDIUM medium = new STGMEDIUM();
            try
            {
                //Define FileGroupDescriptorW format
                FORMATETC format = new FORMATETC();
                format.cfFormat = (short)GetClipboardFormat("FileGroupDescriptorW");
                format.dwAspect = DVASPECT.DVASPECT_CONTENT;
                format.lindex = -1;
                format.ptd = IntPtr.Zero;
                format.tymed = TYMED.TYMED_ISTREAM | TYMED.TYMED_ISTORAGE | TYMED.TYMED_HGLOBAL;

                //Query if format exists in data
                if (data.QueryGetData(format) != NativeMethods.S_OK)
                {
                    log.Debug("No filenames found");
                    return null;
                }

                //Get data into medium
                int retVal = data.GetData(format, out medium);
                if (retVal != NativeMethods.S_OK)
                    throw new Exception(string.Format("Could not get FileGroupDescriptorW format.  Error returned: {0}", retVal));

                //Read medium into byte array
                log.Debug("Reading structure into memory stream");
                byte[] bytes;
                using (MemoryStream stream = new MemoryStream())
                {
                    DataObjectHelper.ReadMediumIntoStream(medium, stream);
                    bytes = new byte[stream.Length];
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.Read(bytes, 0, bytes.Length);
                }
                
                //Copy byte array into unmanaged memory
                log.Debug("Copying structure into unmanaged memory");
                ptrFgd = Marshal.AllocHGlobal(bytes.Length);
                Marshal.Copy(bytes, 0, ptrFgd, bytes.Length);

                //Marshal unmanaged memory to a FILEGROUPDESCRIPTORW struct
                log.Debug("Marshaling unmanaged memory into FILEGROUPDESCRIPTORW struct");
                NativeMethods.FILEGROUPDESCRIPTORW fgd = (NativeMethods.FILEGROUPDESCRIPTORW)Marshal.PtrToStructure(ptrFgd, typeof(NativeMethods.FILEGROUPDESCRIPTORW));
                log.Debug(string.Format("Files found: {0}", fgd.cItems));

                //Create an array to store file names
                string[] filenames = new string[fgd.cItems];

                //Get the pointer to the first file descriptor
                IntPtr ptrFd = IntPtr.Add(ptrFgd, sizeof(uint));

                //Loop for the number of files acording to the file group descriptor
                for (int fdIndex = 0; fdIndex < fgd.cItems; fdIndex++)
                {
                    log.DebugFormat("Getting filename {0}", fdIndex);

                    //Marshal pointer to a FILEDESCRIPTORW struct
                    NativeMethods.FILEDESCRIPTORW fd = (NativeMethods.FILEDESCRIPTORW)Marshal.PtrToStructure(ptrFd, typeof(NativeMethods.FILEDESCRIPTORW));

                    //Get filename of file descriptor and put in array
                    filenames[fdIndex] = fd.cFileName;

                    //Move the file descriptor pointer to the next file descriptor
                    ptrFd = IntPtr.Add(ptrFd, Marshal.SizeOf(fd));
                }

                log.DebugFormat("Filenames found: {0}", string.Join(", ", filenames));
                return filenames;

            }
            finally
            {
                //Release all unmanaged objects
                Marshal.FreeHGlobal(ptrFgd);
                if (medium.pUnkForRelease == null)
                    NativeMethods.ReleaseStgMedium(ref medium);
            }
        }

        internal static void ReadFileContents(NativeMethods.IDataObject data, int index, Stream stream)
        {
            STGMEDIUM medium = new STGMEDIUM();            
            try
            {
                //Define FileContents format
                FORMATETC format = new FORMATETC();
                format.cfFormat = (short)GetClipboardFormat("FileContents");
                format.dwAspect = DVASPECT.DVASPECT_CONTENT;
                format.lindex = index;
                format.ptd = IntPtr.Zero;
                format.tymed = TYMED.TYMED_ISTREAM | TYMED.TYMED_ISTORAGE | TYMED.TYMED_HGLOBAL;

                //Get data
                int retVal = data.GetData(format, out medium);
                if (retVal != NativeMethods.S_OK)
                    throw new Exception(string.Format("Could not get FileContents format.  Error returned: {0}", retVal));

                //Read medium into stream
                ReadMediumIntoStream(medium, stream);
            }
            finally
            {
                //Release all unmanaged objects
                if (medium.pUnkForRelease == null)
                    NativeMethods.ReleaseStgMedium(ref medium);
            }
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

            NativeMethods.ILockBytes iLockBytes = null;
            NativeMethods.IStorage iStorageNew = null;
            IntPtr ptrRead = IntPtr.Zero;
            try
            {
                //Marshal pointer to an IStorage object
                NativeMethods.IStorage iStorage = (NativeMethods.IStorage)Marshal.GetObjectForIUnknown(handle);

                //Create an ILockBytes object on a HGlobal, then create a IStorage object on top of the ILockBytes object
                iLockBytes = NativeMethods.CreateILockBytesOnHGlobal(IntPtr.Zero, true);
                iStorageNew = NativeMethods.StgCreateDocfileOnILockBytes(iLockBytes, 0x00001012, 0);

                //Copy the IStorage object into the new IStorage object
                iStorage.CopyTo(0, null, IntPtr.Zero, iStorageNew);
                iLockBytes.Flush();
                iStorageNew.Commit(0);

                //Get length of ILockBytes byte array
                System.Runtime.InteropServices.ComTypes.STATSTG stat = new System.Runtime.InteropServices.ComTypes.STATSTG();
                iLockBytes.Stat(out stat, 1);
                long length = stat.cbSize;

                //Read bytes into stream
                ptrRead = Marshal.AllocCoTaskMem(sizeof(int));
                byte[] buffer = new byte[4096];     //4 KB buffer
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
            }
            finally
            {
                //Release all unmanaged objects
                Marshal.FreeCoTaskMem(ptrRead);
                if (iStorageNew != null)
                    Marshal.ReleaseComObject(iStorageNew);
                if (iLockBytes != null)
                    Marshal.ReleaseComObject(iLockBytes);
            }
        }

        private static void ReadIStreamIntoStream(IntPtr handle, Stream stream)
        {
            IntPtr ptrRead = IntPtr.Zero;
            try
            {
                //Marshal pointer to an IStream object
                IStream iStream = (IStream)Marshal.GetObjectForIUnknown(handle);

                //Create pointer to integer that stores # of bytes read
                ptrRead = Marshal.AllocCoTaskMem(sizeof(int));

                //Copy IStream into managed stream in chunks
                byte[] buffer = new byte[4096];     //4 KB buffer
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
            }
            finally
            {
                //Release all unmanaged objects
                Marshal.FreeCoTaskMem(ptrRead);
            }

        }

        private static void ReadHGlobalIntoStream(IntPtr handle, Stream stream)
        {
            //Lock HGlobal so it cannot be moved or discarded
            IntPtr source = NativeMethods.GlobalLock(handle);

            if (source == IntPtr.Zero)
                throw new Exception(string.Format("Unable to lock hglobal {0}", source.ToString()));
            try
            {
                //Get size of HGlobal
                int length = NativeMethods.GlobalSize(handle);

                //Copy HGlobal into managed stream in chunks
                byte[] buffer = new byte[4096];     //4 KB buffer
                int bytesToCopy;
                for(int offset = 0; offset < length; offset += buffer.Length)
                {
                    //Copy buffer length or remaining length, whichever is smaller
                    bytesToCopy = Math.Min(buffer.Length, length - offset);
                    Marshal.Copy(source, buffer, 0, bytesToCopy);
                    stream.Write(buffer, 0, bytesToCopy);
                }
            }
            finally
            {
                //Release all unmanaged objects
                NativeMethods.GlobalUnlock(handle);
            }
        }
    }
}
