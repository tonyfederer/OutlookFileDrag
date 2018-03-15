using System.Runtime.InteropServices.ComTypes;

namespace OutlookFileDrag
{
    class FormatEtcEnumerator : IEnumFORMATETC 
    {
        private FORMATETC[] formats;
        private int index = 0;

        public FormatEtcEnumerator(FORMATETC[] formats)
        {
            this.formats = formats;
        }

        public void Clone(out IEnumFORMATETC newEnum)
        {
            //Create new enumerators
            newEnum = new FormatEtcEnumerator(formats);
        }

        public int Next(int celt, FORMATETC[] rgelt, int[] pceltFetched)
        {
            //Fetch number of requested formats
            int fetchCount = 0;
            for (int i = 0; i < celt; i++)
            {
                //If index is past end of formats, stop
                if (index > formats.Length - 1)
                    break;

                //Set format
                rgelt[i] = formats[index];
                fetchCount++;
                index++;
            }

            //Set number of formats fetched
            if (pceltFetched != null && pceltFetched.Length > 0)
                pceltFetched[0] = fetchCount;

            //Return S_OK if all requested formats were returned; otherwise, return S_FALSE
            return (fetchCount == celt ? NativeMethods.S_OK : NativeMethods.S_FALSE);
        }

        public int Reset()
        {
            //Set format index back to 0
            index = 0;
            return NativeMethods.S_OK;
        }

        public int Skip(int celt)
        {
            //Check if incremented index is past end of formats
            if (index + celt > formats.Length - 1)
                return NativeMethods.S_FALSE;
            else
            {
                //Increment index and return
                index += celt;
                return NativeMethods.S_OK;
            }
        }
    }
}
