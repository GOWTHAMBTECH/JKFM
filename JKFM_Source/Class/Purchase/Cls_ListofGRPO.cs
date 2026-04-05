using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JKFM_Source
{
    class Cls_ListofGRPO
    {
        #region Declaration
        SAPbouiCOM.Form objform, oForm;
        SAPbouiCOM.Matrix oMatrix;
        public string TokenNo;
        ClsSBO objSBOAPI;
        #endregion

        #region Constructor
        public Cls_ListofGRPO(ClsSBO objSBO)
        {
            objSBOAPI = objSBO;
        }
        #endregion

        #region Item Event
        public void itemevent(string formuid, ref SAPbouiCOM.ItemEvent pval, ref bool bubbleevent)
        {
            try
            {
                if (pval.BeforeAction == true)
                {
                    objform = objSBOAPI.SBO_Appln.Forms.GetForm(pval.FormTypeEx, pval.FormTypeCount);

                    switch (pval.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            if (pval.ItemUID == "1")
                            {
                                if (objSBOAPI.obj_APInvoice.FormType == "141")
                                {
                                    Get_TokenNo();
                                    if (TokenNo == "")
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("No Token Number found for this GRPO Number", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    }
                                    else
                                    {
                                        objSBOAPI.obj_APInvoice.Lab_Grid_Loading(TokenNo, "", "");
                                        objSBOAPI.obj_APInvoice.FormType = "";
                                    }
                                }
                            }
                            break;

                        case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                            if (pval.ItemUID == "7" && pval.Row != 0)
                            {
                                if (objSBOAPI.obj_APInvoice.FormType == "141")
                                {
                                    Get_TokenNo();
                                    if (TokenNo == "")
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("No Token Number found for this GRPO Number", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    }
                                    else
                                    {
                                        objSBOAPI.obj_APInvoice.Lab_Grid_Loading(TokenNo, "", "");
                                        objSBOAPI.obj_APInvoice.FormType = "";
                                    }
                                }
                            }
                            break;
                    }
                }
                else
                {
                    switch (pval.EventType)
                    {

                    }
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Menu Event
        public void MenuEvent(ref SAPbouiCOM.MenuEvent pval, ref bool BubbleEvent)
        {
            try
            {
                if (pval.BeforeAction == true)
                {
                    switch (pval.MenuUID)
                    {

                    }
                }
                else if (pval.BeforeAction == false)
                {
                    switch (pval.MenuUID)
                    {

                    }
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region Get TokenNo
        public void Get_TokenNo()
        {
            try
            {
                oMatrix = objform.Items.Item("7").Specific;
                for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                {
                    if (oMatrix.IsRowSelected(i) == true)
                    {
                        TokenNo = oMatrix.Columns.Item("U_AVA_PURCHASETOKENN").Cells.Item(i).Specific.Value.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion
    }
}
