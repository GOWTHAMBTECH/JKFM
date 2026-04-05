using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using AvanikoFramework;

namespace JKFM_Source
{
    public class SubMain
    {

        #region Declaration
        public ClsMain ObjBP;
        #endregion

        #region Main
        public static void Main()
        {
            ClsMain ObjBP = new ClsMain();
            if (ObjBP.Initialise())
            {
                ObjBP.objSBOAPI.SBO_Appln.StatusBar.SetText("JKFM Add-On is Connected Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                System.Windows.Forms.Application.Run();
            }
            else
            {
                SAPbouiCOM.Form oForm = null;
                oForm = ObjBP.objSBOAPI.SBO_Appln.Forms.ActiveForm;
                if (oForm.TypeEx == "AV_ADLUF")
                {
                    System.Windows.Forms.Application.Run();
                }
                else
                {
                    ObjBP.objSBOAPI.SBO_Appln.StatusBar.SetText("JKFM Add-On Connection Failed", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    System.Windows.Forms.Application.Exit();
                }
            }
        }
        #endregion

    }
}
