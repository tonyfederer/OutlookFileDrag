using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using log4net;

namespace OutlookFileDrag
{
    //Class that wraps Outlook data object and adds support for CF_HDROP format
    class DataObjectWrapper : NativeMethods.IDataObject
    {
        private NativeMethods.IDataObject innerData;

        public DataObjectWrapper(NativeMethods.IDataObject innerData)
        {
            this.innerData = innerData;
        }

        public int EnumFormatEtc(DATADIR direction, out IEnumFORMATETC ppenumFormatEtc)
        {
            return innerData.EnumFormatEtc(direction, out ppenumFormatEtc);
        }

        public int GetCanonicalFormatEtc(ref FORMATETC formatIn, out FORMATETC formatOut)
        {
            return innerData.GetCanonicalFormatEtc(formatIn, out formatOut);
        }

        public int GetData(ref FORMATETC format, out STGMEDIUM medium)
        {
            return innerData.GetData(format, out medium);
        }

        public int GetDataHere(ref FORMATETC format, ref STGMEDIUM medium)
        {
            return innerData.GetDataHere(format, medium);
        }

        public int QueryGetData(ref FORMATETC format)
        {
            return innerData.QueryGetData(format);
        } 

        public int SetData(ref FORMATETC formatIn, ref STGMEDIUM medium, bool release)
        {
            return innerData.SetData(formatIn, medium, release);
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

    }
}
