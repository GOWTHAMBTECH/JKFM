using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JKFM_Source
{
    class Cls_QueryManagerUpdate
    {

        #region Declaration
        SAPbouiCOM.Form objform;
        ClsSBO objSBOAPI;
        #endregion        

        #region Constructor
        public Cls_QueryManagerUpdate(ClsSBO objSBO)
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
                    //switch (pval.EventType)
                    //{
                    //}
                }
                else
                {
                    switch (pval.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                            if (pval.ItemUID == "Item_1" && pval.FormMode == Convert.ToInt32(SAPbouiCOM.BoFormMode.fm_UPDATE_MODE))
                            {
                                DialogResult Dialog;
                                Dialog = (DialogResult)objSBOAPI.SBO_Appln.MessageBox("Query Manager will be updated.Do you wish to Proceed", 1, "Yes", "No");
                                if (Dialog == DialogResult.OK)
                                {                                    
                                    objSBOAPI.SBO_Appln.StatusBar.SetText("Creating Query Categories and Queries Please Wait...........", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                    SAPbobsCOM.Recordset Orec = null;
                                    Orec = (SAPbobsCOM.Recordset)objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    string Str = "";
                                    SAPbouiCOM.Grid oGrid;
                                    oGrid = (SAPbouiCOM.Grid)objform.Items.Item("Item_0").Specific;
                                    SAPbouiCOM.CheckBoxColumn oChk;
                                    oChk = (SAPbouiCOM.CheckBoxColumn)oGrid.Columns.Item("Select");

                                    for (int i = 0; i <= oGrid.Rows.Count - 1; i++)
                                    {
                                        if (oChk.IsChecked(i) == true)
                                        {
                                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                            {
                                                Str = "UPDATE OUQR SET \"U_AV_Sel\" = 'Y' WHERE \"QName\" = '" + oGrid.DataTable.GetValue("QueryName", i) + "' AND \"QCategory\" = (SELECT A.\"CategoryId\" FROM OQCN A WHERE A.\"CatName\" = 'AddonQuery_Hana')";
                                            }
                                            else
                                            {
                                                Str = "UPDATE OUQR SET U_AV_Sel = 'Y' WHERE QName = '" + oGrid.DataTable.GetValue("QueryName", i) + "' AND QCategory = (SELECT A.CategoryId FROM OQCN A WHERE A.CatName = 'AddonQuery_SQL')";
                                            }
                                                                                        
                                        }
                                        else
                                        {
                                            if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                            {
                                                Str = "UPDATE OUQR SET \"U_AV_Sel\" = 'N' WHERE \"QName\" = '" + oGrid.DataTable.GetValue("QueryName", i) + "' AND \"QCategory\" = (SELECT A.\"CategoryId\" FROM OQCN A WHERE A.\"CatName\" = 'AddonQuery_Hana')";
                                            }
                                            else
                                            {
                                                Str = "UPDATE OUQR SET U_AV_Sel = 'N' WHERE QName = '" + oGrid.DataTable.GetValue("QueryName", i) + "' AND QCategory = (SELECT A.CategoryId FROM OQCN A WHERE A.CatName = 'AddonQuery_SQL')";
                                            }
                                        }
                                        Orec.DoQuery(Str);
                                    }
                                    objSBOAPI.objMain.CreateQueries();                           
                                    objform.Close();
                                }
                            }
                            break;
                        case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                            if (pval.ItemUID == "Item_0" && pval.ColUID == "Select" && pval.Row == -1)
                            {
                                SAPbouiCOM.DataTable Dt = objform.DataSources.DataTables.Item("DT_0");
                                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)objform.Items.Item("Item_0").Specific;
                                objform.Freeze(true);
                                if (oGrid.DataTable.IsEmpty == false)
                                {
                                    for (int i = 0; i < Dt.Rows.Count; i++)
                                    {
                                        if (Dt.GetValue("Select", i) == "Y")
                                        {
                                            Dt.SetValue("Select", i, "N");
                                        }
                                        else
                                        {
                                            Dt.SetValue("Select", i, "Y");
                                        }
                                    }
                                }
                                objform.Freeze(false);
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
                    if (pval.MenuUID == "")
                    {

                    }
                    else if (pval.MenuUID == "1281")
                    {

                    }
                    else if (pval.MenuUID == "1282")
                    {

                    }
                    else if ((pval.MenuUID == "1288") || (pval.MenuUID == "1289") || (pval.MenuUID == "1290") || (pval.MenuUID == "1291"))
                    {

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

    }
}
