using System;
using EasyHook;
using log4net;

namespace OutlookFileDrag
{
    class DragDropHook : IDisposable
    {
        private static ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private LocalHook hook;

        public void StartHook()
        {
            //Hook OLE drag and drop event
            log.Info("Hooking DoDragDrop method of ole32.dll");
            hook = EasyHook.LocalHook.Create(EasyHook.LocalHook.GetProcAddress("ole32.dll", "DoDragDrop"),
                new NativeMethods.DragDropDelegate(DragDropHook.DoDragDropHook), null);

            //Only hook this thread (threadId == 0 == GetCurrentThreadId)
            hook.ThreadACL.SetInclusiveACL(new Int32[] { 0 });
            log.Info("Hooked DoDragDrop method");
        }

        public static int DoDragDropHook(NativeMethods.IDataObject pDataObj, NativeMethods.IDropSource pDropSource, uint dwOKEffects, uint[] pdwEffect)
        {
            try
            {
                log.Debug("Drag started");
                if (!DataObjectHelper.GetDataPresent(pDataObj, "FileGroupDescriptorW"))
                {
                    log.Debug("No virtual files found -- continuing original drag");
                    return NativeMethods.DoDragDrop(pDataObj, pDropSource, dwOKEffects, pdwEffect);
                }

                //Start new drag
                log.Debug("Virtual files found -- starting new drag adding CF_HDROP and CFSTR_SHELLIDLIST formsts");
                log.DebugFormat("Files: {0}", string.Join(",", DataObjectHelper.GetFilenames(pDataObj)));

                NativeMethods.IDataObject newDataObj = new OutlookDataObject(pDataObj);
                int result = NativeMethods.DoDragDrop(newDataObj, pDropSource, dwOKEffects, pdwEffect);

                //Get result
                log.DebugFormat("DoDragDrop result: {0}", result);
                return result;
            }
            catch (Exception ex)
            {
                log.Warn("Dragging error", ex);
                return NativeMethods.DRAGDROP_S_CANCEL;
            }
        }

        public void Dispose()
        {
            //Dispose hook
            if (hook != null)
            {
                log.Info("Disposing hook");
                hook.Dispose();
                hook = null;
            }
        }
    }

}
