using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Media;

namespace JKFM_Source
{
    class Cls_ShiftMaster
    {
        #region Declaration
        SAPbouiCOM.Form objform;
        ClsSBO objSBOAPI;

        SAPbouiCOM.DBDataSource oDBDSHeader;
        string UDOID;
        #endregion        

        #region Constructor
        public Cls_ShiftMaster(ClsSBO objSBO)
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
                        case BoEventTypes.et_CLICK:
                            switch(pval.ItemUID)
                            {
                                case "1":
                                    if ((objform.Mode == BoFormMode.fm_ADD_MODE || objform.Mode == BoFormMode.fm_UPDATE_MODE) && !ValidateAll())
                                    {
                                        SystemSounds.Asterisk.Play();
                                        bubbleevent = false;
                                    }
                                    break;
                            }
                            
                            break;
                    }
                }
                else
                {
                    switch (pval.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            switch(pval.ItemUID)
                            {
                                case "1":
                                    if (pval.ActionSuccess == true && objform.Mode == BoFormMode.fm_ADD_MODE)
                                    {
                                        InitForm();
                                    }
                                    break;
                            }
                            break;

                        case BoEventTypes.et_VALIDATE:
                            switch(pval.ItemUID)
                            {
                                case "t_shrs":
                                    string Shift_Start_Time = Convert.ToString(objform.Items.Item("t_stime").Specific.Value);
                                    string Start_Break_Time = Convert.ToString(objform.Items.Item("t_bstime").Specific.Value);
                                    string End_Break_Time = Convert.ToString(objform.Items.Item("t_betime").Specific.Value);
                                    Convert.ToString(objform.Items.Item("t_etime").Specific.Value);
                                    string Shift_Hours = Convert.ToString(objform.Items.Item("t_shrs").Specific.Value);
                                    CheckBox checkBox = (CheckBox)objform.Items.Item("c_ILun").Specific;
                                    string LunchEnable = checkBox.Checked.ToString();
                                    
                                    if (Shift_Start_Time != string.Empty && Start_Break_Time != string.Empty && End_Break_Time != string.Empty && Shift_Hours != string.Empty)
                                    {
                                        int Shift_Hours_Num = Convert.ToInt32(objform.Items.Item("t_shrs").Specific.Value);
                                        string QueryStr = null;
                                        if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                        {
                                            QueryStr = "CALL \"@AIS_SalePlanning_ShiftCalculateEndTime\"('" + Shift_Start_Time + "','" + Start_Break_Time + "','" + End_Break_Time + "'," + Convert.ToString(Shift_Hours_Num) + ",'" + LunchEnable + "')";
                                        }
                                        else
                                        {
                                            QueryStr = "Exec [dbo].[@AIS_SalePlanning_ShiftCalculateEndTime]'" + Shift_Start_Time + "','" + Start_Break_Time + "','" + End_Break_Time + "'," + Convert.ToString(Shift_Hours_Num) + ",'" + LunchEnable + "'";
                                        }
                                            
                                        Recordset Orec = (Recordset)objSBOAPI.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        Orec.DoQuery(QueryStr);
                                        if (Orec.RecordCount > 0)
                                        {
                                            Orec.MoveFirst();
                                            string ToTime = Orec.Fields.Item("ToTime").Value.ToString().Trim();
                                            ToTime = ToTime.Replace(":", "");
                                            oDBDSHeader.SetValue("U_SEndTime", 0, ToTime);
                                            oDBDSHeader.SetValue("U_WorkHours", 0, Orec.Fields.Item("WorkHours").Value.ToString().Trim());
                                        }
                                    }
                                    break;
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
                }
                else if (pval.BeforeAction == false)
                {
                    switch (pval.MenuUID)
                    {
                        case "OSFT":
                            objform = objSBOAPI.LoadForm("ShiftMaster.xml", "OSFT");
                            objform = objSBOAPI.SBO_Appln.Forms.Item("OSFT");
                            oDBDSHeader = objform.DataSources.DBDataSources.Item("@AIS_OSFT");
                            objform.Mode = BoFormMode.fm_ADD_MODE;
                            DefineModesForFields();
                            InitForm();
                            break;

                        case "1282":
                            InitForm();
                            objform.ActiveItem = "t_code";
                            break;

                        case "1281":
                            if (objform.Mode != BoFormMode.fm_FIND_MODE)
                                objform.Mode = BoFormMode.fm_FIND_MODE;
                            objform.ActiveItem = "t_code";
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.MessageBox(ex.Message);
            }
        }
        #endregion

        #region InitForm
        public void InitForm()
        {
            try
            {
                objform.Freeze(true);
                objform.ActiveItem = "t_code";
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Init Form Method Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                objform.Freeze(false);
            }
        }
        #endregion

        #region DefineModesForFields
        public void DefineModesForFields()
        {
            try
            {
                objform.Items.Item("t_code").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True);
                objform.Items.Item("t_etime").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True);
                objform.Items.Item("t_whrs").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 4, BoModeVisualBehavior.mvb_True);
                objform.Items.Item("t_code").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 1, BoModeVisualBehavior.mvb_False);
                objform.Items.Item("t_code").SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, 2, BoModeVisualBehavior.mvb_True);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Define Modes For Fields Method Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #region ValidateAll
        public bool ValidateAll()
        {
            bool flag;
            try
            {
                flag = true;
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                objSBOAPI.SBO_Appln.StatusBar.SetText("Validate Function Failed:" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                flag = false;
            }
            return flag;
        }
        #endregion

    }
}
