using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using log4net;

namespace OutlookFileDrag
{
    //Class that wraps Outlook data object and adds support for CF_HDROP format
    class OutlookDataObject : NativeMethods.IDataObject, ICustomQueryInterface  
    {
        private NativeMethods.IDataObject innerData;
        private string[] tempFilenames;
        private static ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public bool FilesDropped { get; private set; }

        public OutlookDataObject(NativeMethods.IDataObject innerData)
        {
            this.innerData = innerData;
        }

        public int EnumFormatEtc(DATADIR direction, out IEnumFORMATETC ppenumFormatEtc)
        {
            IEnumFORMATETC origEnum = null;
            try
            {
                log.DebugFormat("IDataObject.EnumFormatEtc called -- direction {0}", direction);
                switch (direction)
                {
                    case DATADIR.DATADIR_GET:
                        //Get original enumerator
                        int result = innerData.EnumFormatEtc(direction, out origEnum);
                        if (result != NativeMethods.S_OK)
                        {
                            ppenumFormatEtc = null;
                            return result;
                        }

                        //Enumerate original formats
                        List<FORMATETC> formats = new List<FORMATETC>();
                        FORMATETC[] buffer = new FORMATETC[] { new FORMATETC() };
                        while (origEnum.Next(1, buffer, null) == NativeMethods.S_OK)
                        {
                            //Convert format from short to unsigned short
                            ushort cfFormat = (ushort) buffer[0].cfFormat;

                            //Do not return text formats -- some applications try to get text before files
                            if (cfFormat != NativeMethods.CF_TEXT && cfFormat != NativeMethods.CF_UNICODETEXT && cfFormat != (ushort)DataObjectHelper.GetClipboardFormat("Csv"))
                                formats.Add(buffer[0]);
                        }

                        //Add CF_HDROP format
                        FORMATETC format = new FORMATETC();
                        format.cfFormat = NativeMethods.CF_HDROP;
                        format.dwAspect = DVASPECT.DVASPECT_CONTENT;
                        format.lindex = -1;
                        format.ptd = IntPtr.Zero;
                        format.tymed = TYMED.TYMED_HGLOBAL;
                        formats.Add(format);

                        //Return new enumerator for available formats
                        ppenumFormatEtc = new FormatEtcEnumerator(formats.ToArray());
                        return NativeMethods.S_OK;

                    case DATADIR.DATADIR_SET:
                        //Return original enumerator
                        return innerData.EnumFormatEtc(direction, out ppenumFormatEtc);
                    default:
                        //Invalid direction
                        ppenumFormatEtc = null;
                        return NativeMethods.E_INVALIDARG;
                }

            }
            catch (Exception ex)
            {
                log.Error("Exception in IDataObject.EnumFormatEtc", ex);
                ppenumFormatEtc = null;
                return NativeMethods.E_UNEXPECTED;
            }
            finally
            {
                //Release all unmanaged objects
                if (origEnum != null)
                    Marshal.ReleaseComObject(origEnum);
            }
        }

        public int GetCanonicalFormatEtc(ref FORMATETC formatIn, out FORMATETC formatOut)
        {
            try
            {
                log.DebugFormat("IDataObject.GetCanonicalFormatEtc called -- cfFormat {0} dwAspect {1} lindex {2} ptd {3} tymed {4}", formatIn.cfFormat, formatIn.dwAspect, formatIn.lindex, formatIn.ptd, formatIn.tymed);
                if (formatIn.cfFormat == NativeMethods.CF_HDROP)
                {
                    //Copy input format to output format
                    formatOut = new FORMATETC();
                    formatOut.cfFormat = formatIn.cfFormat;
                    formatOut.dwAspect = formatIn.dwAspect;
                    formatOut.lindex = formatIn.lindex;
                    formatOut.ptd = IntPtr.Zero;
                    formatOut.tymed = formatIn.tymed;
                    
                    return NativeMethods.DATA_S_SAMEFORMATETC;
                }
                else
                    return innerData.GetCanonicalFormatEtc(formatIn, out formatOut);
            }
            catch (Exception ex)
            {
                log.Error("Exception in IDataObject.GetCanonicalFormatEtc", ex);
                formatOut = new FORMATETC();
                return NativeMethods.E_UNEXPECTED;
            }
        }

        public int GetData(ref FORMATETC format, out STGMEDIUM medium)
        {
            try
            {
                //Get data into passed medium
                log.DebugFormat("IDataObject.GetData called -- cfFormat {0} dwAspect {1} lindex {2} ptd {3} tymed {4}", format.cfFormat, format.dwAspect, format.lindex, format.ptd, format.tymed);
                log.DebugFormat("Format name: {0}", System.Windows.Forms.DataFormats.GetFormat((ushort)format.cfFormat).Name);

                if (format.cfFormat == NativeMethods.CF_HDROP)
                {
                    medium = new STGMEDIUM();

                    //Validate index
                    if (format.lindex != -1)
                        return NativeMethods.DV_E_LINDEX;
                    //Validate medium type
                    if (!format.tymed.HasFlag(TYMED.TYMED_HGLOBAL))
                        return NativeMethods.DV_E_TYMED;
                    //Validate DV aspect
                    if (format.dwAspect != DVASPECT.DVASPECT_CONTENT)
                        return NativeMethods.DV_E_DVASPECT;

                    //Extract files if not already extracted
                    if (tempFilenames == null)
                        ExtractFiles();

                    //Get list of dropped files
                    log.Debug("Setting drop files");
                    DataObjectHelper.SetDropFiles(ref medium, tempFilenames);
                    FilesDropped = true;
                    return NativeMethods.S_OK;
                }
                else if (format.cfFormat == NativeMethods.CF_TEXT || format.cfFormat == NativeMethods.CF_UNICODETEXT || format.cfFormat == (ushort)DataObjectHelper.GetClipboardFormat("Csv"))
                {
                    //Do not return text formats -- some applications try to get text before files
                    medium = new STGMEDIUM();
                    return NativeMethods.DV_E_FORMATETC;
                }
                else
                {
                    int result =  innerData.GetData(format, out medium);
                    log.DebugFormat("Result: {0}", result);
                    return result;
                }

            }
            catch (Exception ex)
            {
                log.Error("Exception in IDataObject.GetData", ex);
                medium = new STGMEDIUM();
                return NativeMethods.E_UNEXPECTED;
            }
        }

        public int GetDataHere(ref FORMATETC format, ref STGMEDIUM medium)
        {
            log.DebugFormat("IDataObject.QueryGetData called -- cfFormat {0} dwAspect {1} lindex {2} ptd {3} tymed {4}", format.cfFormat, format.dwAspect, format.lindex, format.ptd, format.tymed);
            return NativeMethods.E_NOTIMPL;
        }

        public int QueryGetData(ref FORMATETC format)
        {
            try
            {
                log.DebugFormat("IDataObject.QueryGetData called -- cfFormat {0} dwAspect {1} lindex {2} ptd {3} tymed {4}", format.cfFormat, format.dwAspect, format.lindex, format.ptd, format.tymed);
                log.DebugFormat("Format name: {0}", System.Windows.Forms.DataFormats.GetFormat((ushort)format.cfFormat).Name);
                if (format.cfFormat == NativeMethods.CF_HDROP)
                {
                    //Validate index
                    if (format.lindex != -1)
                        return NativeMethods.DV_E_LINDEX;
                    //Validate medium type
                    if (!format.tymed.HasFlag(TYMED.TYMED_HGLOBAL))
                        return NativeMethods.DV_E_TYMED;
                    //Validate DV aspect
                    if (format.dwAspect != DVASPECT.DVASPECT_CONTENT)
                        return NativeMethods.DV_E_DVASPECT;

                    log.DebugFormat("IDataObject.QueryGetData result: {0}", NativeMethods.S_OK);
                    return NativeMethods.S_OK;
                }
                else if (format.cfFormat == NativeMethods.CF_TEXT || format.cfFormat == NativeMethods.CF_UNICODETEXT || format.cfFormat == (ushort)DataObjectHelper.GetClipboardFormat("Csv"))
                {
                    //Do not return text formats -- some applications try to get text before files
                    return NativeMethods.DV_E_FORMATETC;
                }
                else
                {
                    int result = innerData.QueryGetData(format);
                    log.DebugFormat("Result: {0}", result);
                    return result;
                }

            }
            catch (Exception ex)
            {
                log.Error("Exception in IDataObject.QueryGetData", ex);
                return NativeMethods.E_UNEXPECTED;
            }
        }

        public int SetData(ref FORMATETC formatIn, ref STGMEDIUM medium, bool release)
        {
            try
            {
                log.DebugFormat("IDataObject.SetData called -- cfFormat {0} dwAspect {1} lindex {2} ptd {3} tymed {4}", formatIn.cfFormat, formatIn.dwAspect, formatIn.lindex, formatIn.ptd, formatIn.tymed);
                log.DebugFormat("Format name: {0}", System.Windows.Forms.DataFormats.GetFormat((ushort)formatIn.cfFormat).Name);
                int result = innerData.SetData(formatIn, medium, release);
                log.DebugFormat("Result: {0}", result);
                return result;

            }
            catch (Exception ex)
            {
                log.Error("Exception in IDataObject.SetData", ex);
                return NativeMethods.E_UNEXPECTED;
            }
        }

        public int DAdvise(ref FORMATETC pFormatetc, ADVF advf, IAdviseSink adviseSink, out int connection)
        {
            return innerData.DAdvise(pFormatetc, advf, adviseSink, out connection);
        }

        public int DUnadvise(int connection)
        {
            return innerData.DUnadvise(connection);
        }

        public int EnumDAdvise(out IEnumSTATDATA enumAdvise)
        {
            return innerData.EnumDAdvise(out enumAdvise);
        }

        private void ExtractFiles()
        {
            //Get filenames being dragged
            log.Debug("Getting filenames");
            string[] filenames = DataObjectHelper.GetFilenames(this.innerData);
            log.DebugFormat("Filenames: {0}", string.Join(",", filenames));

            //Get temporary folder
            log.Debug("Creating temp folder");
            string tempPath = FileUtility.GetTempPath();
            log.DebugFormat("Temp folder: {0}", tempPath);

            //Save files to temporary directory
            tempFilenames = new string[filenames.Length];
            for (int fileIndex = 0; fileIndex < filenames.Length; fileIndex++)
            {
                tempFilenames[fileIndex] = FileUtility.GetUniqueFilename(Path.Combine(tempPath, filenames[fileIndex]));
                log.DebugFormat("Extracting file {0}", filenames[fileIndex]);
                using (FileStream fs = new FileStream(tempFilenames[fileIndex], FileMode.Create))
                {
                    DataObjectHelper.ReadFileContents(this.innerData, fileIndex, fs);
                }
            }
        }

        public CustomQueryInterfaceResult GetInterface(ref Guid iid, out IntPtr ppv)
        {
            ppv = IntPtr.Zero;
            try
            {
                log.DebugFormat("Get COM interface {0}", iid);

                //For IDataObject interface, use interface on this object
                if (iid == new Guid("0000010E-0000-0000-C000-000000000046"))
                {
                    log.DebugFormat("Interface handled");
                    return CustomQueryInterfaceResult.NotHandled;
                }

                else
                {
                    //For all other interfaces, use interface on original object
                    IntPtr pUnk = Marshal.GetIUnknownForObject(this.innerData);
                    int retVal = Marshal.QueryInterface(pUnk, ref iid, out ppv);
                    if (retVal == NativeMethods.S_OK)
                    {
                        log.DebugFormat("Interface handled by inner object");
                        return CustomQueryInterfaceResult.Handled;
                    }
                    else
                    {
                        log.DebugFormat("Interface not handled by inner object");
                        return CustomQueryInterfaceResult.Failed;
                    }
                }

            }
            catch (Exception ex)
            {
                log.Error("Exception in ICustomQueryInterface", ex);
                return CustomQueryInterfaceResult.Failed;
            }

        }
    }
}
