using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OutlookFileDrag
{
    class DropSourceWrapper : NativeMethods.IDropSource
    {
        private NativeMethods.IDropSource innerSource;

        public DropSourceWrapper(NativeMethods.IDropSource innerSource)
        {
            this.innerSource = innerSource;
        }
        
        public int QueryContinueDrag(bool fEscapePressed, int grfKeyState)
        {
            return innerSource.QueryContinueDrag(fEscapePressed, grfKeyState);
        }

        public int GiveFeedback(int dwEffect)
        {
            return innerSource.GiveFeedback(dwEffect);
        }
    }
}
