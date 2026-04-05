using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JKFM_Source
{
    class Cls_GRPO
    {
        #region Declaration
        SAPbouiCOM.Form objform;
        SAPbouiCOM.ComboBox oCombo;
        string TokenNo;
        ClsSBO objSBOAPI;
        #endregion

        #region Constructor
        public Cls_GRPO(ClsSBO objSBO)
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

                    }
                }
                else
                {
                    switch (pval.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:

                            break;

                        case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                            switch (pval.ItemUID)
                            {
                                case "10000329":

                                    oCombo = objform.Items.Item("10000329").Specific;
                                    if (oCombo.Selected.Value == "A/P Invoice")
                                    {
                                        TokenNo = objform.DataSources.DBDataSources.Item("OPDN").GetValue("U_AVA_PURCHASETOKENNO", 0).ToString();
                                        if (TokenNo == "")
                                        {
                                            objSBOAPI.SBO_Appln.StatusBar.SetText("No Token Number found for this GRPO Number", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        }
                                        else
                                        {
                                            objSBOAPI.obj_APInvoice.Lab_Grid_Loading(TokenNo, "", "");
                                        }
                                    }
                                    break;
                            }
                            break;

                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            switch (pval.ItemUID)
                            {

                            }
                            break;
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
    }
}
