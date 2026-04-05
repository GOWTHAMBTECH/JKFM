using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JKFM_Source
{
    class Cls_PickListData
    {
        #region Declaration
        SAPbouiCOM.Form objform, oForm;
        ClsSBO objSBOAPI;
        #endregion        

        #region Constructor
        public Cls_PickListData(ClsSBO objSBO)
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
                        case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                            objform = objSBOAPI.SBO_Appln.Forms.GetForm("60020", -1);
                            break;
                    }
                }
                else
                {
                    switch (pval.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

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
                }
                else if (pval.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.MessageBox(ex.Message);
            }
        }
        #endregion

        #region FORM DATA EVENT
        public void FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            if (objform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                oForm = objSBOAPI.SBO_Appln.Forms.GetForm("81", -1);
                                SAPbobsCOM.Recordset Orec = null;
                                Orec = (SAPbobsCOM.Recordset)objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                string str1 = string.Empty;
                                SAPbouiCOM.ComboBox comboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("c_vehicle").Specific;
                                if (comboBox.Selected != null)
                                    str1 = comboBox.Selected.Description.ToString().Trim();
                                string str2 = objform.Items.Item("11").Specific.Value.ToString().Trim();
                                string strSql = null;
                                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                {
                                    strSql = "Update \"OPKL\" Set \"U_Vehicle\" ='" + str1 + "' Where \"AbsEntry\" ='" + str2 + "'";
                                }
                                else
                                {
                                    strSql = "Update OPKL Set U_Vehicle ='" + str1 + "' Where AbsEntry ='" + str2 + "'";
                                }
                                    
                                Orec.DoQuery(strSql);
                            }                            break;
                    }

                }
            }
            catch (Exception ex)
            {
                objSBOAPI.SBO_Appln.MessageBox(ex.Message);
            }
        }

        #endregion

    }
}
