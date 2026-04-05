using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JKFM_Source
{
    class Cls_InsuranceDetails
    {
        #region Declaration
        SAPbouiCOM.Form objform;
        SAPbouiCOM.ComboBox oCombo;
        ClsSBO objSBOAPI;
        #endregion

        #region Constructor
        public Cls_InsuranceDetails(ClsSBO objSBO)
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

                            if (pval.ItemUID == "1" && pval.FormMode == Convert.ToInt32(SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                            {
                                objform.Items.Item("Et_BalAmt").Specific.Value = Convert.ToDouble(objform.Items.Item("Et_InsAmt").Specific.Value);
                                Validation(ref bubbleevent);
                            }
                            break;
                    }
                }
                else
                {
                    switch (pval.EventType)
                    {

                        case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:

                            if (pval.ItemUID == "Cb_Code" && pval.ItemChanged == true)
                            {
                                oCombo = objform.Items.Item("Cb_Code").Specific;
                                SAPbobsCOM.Recordset oRec = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                string Str = "SELECT \"Name\",\"U_AVA_INSAMT\" FROM \"@AVA_OSIP\" where \"U_AVA_INSPNO\" = '" + oCombo.Selected.Value + "'";
                                oRec.DoQuery(Str);
                                if (oRec.RecordCount > 0)
                                {
                                    objform.Items.Item("Et_Name").Specific.Value = oRec.Fields.Item("Name").Value;
                                    objform.Items.Item("Et_InsAmt").Specific.Value = Convert.ToDouble(oRec.Fields.Item("U_AVA_INSAMT").Value);
                                }
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
                        case "AV_INSDTM":

                            objform = objSBOAPI.LoadForm("InsuranceDetails.xml", "AVA_INSDET");
                            oCombo = objform.Items.Item("Cb_Code").Specific;

                            SAPbobsCOM.Recordset orec;
                            orec = objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string str = "SELECT \"U_AVA_INSPNO\" FROM \"@AVA_OSIP\"";

                            orec.DoQuery(str);
                            if (orec.RecordCount > 0)
                            {
                                while (!orec.EoF)
                                {
                                    oCombo.ValidValues.Add(orec.Fields.Item("U_AVA_INSPNO").Value, orec.Fields.Item("U_AVA_INSPNO").Value);
                                    orec.MoveNext();
                                }
                                oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                            }
                            break;

                        case "1281":

                            objform.Items.Item("Cb_Code").Enabled = true;
                            objform.Items.Item("Et_Date").Enabled = true;
                            objform.Items.Item("Et_Name").Enabled = true;
                            objform.Items.Item("Et_InsAmt").Enabled = true;
                            break;

                        case "1282":

                            objform.Items.Item("Cb_Code").Enabled = true;
                            objform.Items.Item("Et_Date").Enabled = true;
                            objform.Items.Item("Et_InsAmt").Enabled = true;
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
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:

                            objform.Items.Item("Cb_Code").Enabled = false;
                            objform.Items.Item("Et_Date").Enabled = false;
                            objform.Items.Item("Et_InsAmt").Enabled = false;
                            objform.Items.Item("Et_Name").Enabled = false;

                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                objSBOAPI.SBO_Appln.MessageBox(ex.Message);
            }
        }

        #endregion

        #region Validation 
        public void Validation(ref bool Bubbleevent)
        {
            try
            {
                if (objform.Items.Item("Cb_Code").Specific.Value == "")
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Insurance Policy No is Mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    Bubbleevent = false;
                    return;
                }
                else if (Convert.ToDouble(objform.Items.Item("Et_InsAmt").Specific.Value) <= 0)
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Insurance Amount must be greater than zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    Bubbleevent = false;
                    return;
                }
                else if (Convert.ToDouble(objform.Items.Item("Et_BalAmt").Specific.Value) <= 0)
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("Amount must be greater than zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    Bubbleevent = false;
                    return;
                }
                else if (objform.Items.Item("Et_Date").Specific.Value == "")
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("DocDate is Mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    Bubbleevent = false;
                    return;
                }
                else if (objform.Items.Item("Et_FDate").Specific.Value == "")
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("From Date is Mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    Bubbleevent = false;
                    return;
                }
                else if (objform.Items.Item("Et_TDate").Specific.Value == "")
                {
                    objSBOAPI.SBO_Appln.StatusBar.SetText("To Date is Mandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    Bubbleevent = false;
                    return;
                }
                else if (objform.Items.Item("Et_FDate").Specific.Value != "" && objform.Items.Item("Et_TDate").Specific.Value != "")
                {
                    string Fdate = objSBOAPI.GetDateFromField(objform.Items.Item("Et_FDate").Specific.Value);
                    string Tdate = objSBOAPI.GetDateFromField(objform.Items.Item("Et_TDate").Specific.Value);
                    int result = string.Compare(Fdate, Tdate);
                    if (result > 0)
                    {
                        objSBOAPI.SBO_Appln.StatusBar.SetText("From Date should not be greater than To Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        Bubbleevent = false;
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                Bubbleevent = false;
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion
    }
}
