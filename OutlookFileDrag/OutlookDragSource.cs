using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using log4net;

namespace OutlookFileDrag
{
    class OutlookDragSource : NativeMethods.IDropSource, ICustomQueryInterface
    {
        private IntPtr innerData;
        private static ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public OutlookDragSource(IntPtr innerData)
        {
            this.innerData = innerData;
        }


        public CustomQueryInterfaceResult GetInterface(ref Guid iid, out IntPtr ppv)
        {
            ppv = IntPtr.Zero;
            try
            {
                log.DebugFormat("Get COM interface {0}", iid);

                //For all other interfaces, use interface on original object
                //IntPtr pUnk = Marshal.GetIUnknownForObject(this.innerData);
                IntPtr pUnk = this.innerData;
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
            catch (Exception ex)
            {
                log.Error("Exception in ICustomQueryInterface", ex);
                return CustomQueryInterfaceResult.Failed;
            }

        }

        public int QueryContinueDrag([MarshalAs(UnmanagedType.Bool)] bool fEscapePressed, int grfKeyState)
        {
            throw new NotImplementedException();
        }

        public int GiveFeedback(int dwEffect)
        {
            throw new NotImplementedException();
        }
    }
}
