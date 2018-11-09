using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OutlookFileDrag
{
    class OutlookDropSource : NativeMethods.IDropSource
    {
        private int? originalKeyState = null;

        public int QueryContinueDrag(bool fEscapePressed, int grfKeyState)
        {
            //If key state has changed since drag started, finish drop
            if (originalKeyState == null)
                originalKeyState = grfKeyState;
            else if (grfKeyState != originalKeyState)
                return NativeMethods.DRAGDROP_S_DROP;

            //If escape has been pressed, cancel drop, otherwise continue
            if (fEscapePressed)
                return NativeMethods.DRAGDROP_S_CANCEL;
            else
                return NativeMethods.S_OK;

        }

        public int GiveFeedback(int dwEffect)
        {
            return NativeMethods.DRAGDROP_S_USEDEFAULTCURSORS;
        }

    }
}
