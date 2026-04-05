using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JKFM_Source
{
    public class WindowWrapper : System.Windows.Forms.IWin32Window
    {

        private IntPtr _hwnd;

        public WindowWrapper(IntPtr handle)
        {
            _hwnd = handle;
        }

        public System.IntPtr Handle
        {
            get
            {
                return _hwnd;
            }
        }
    }

}
