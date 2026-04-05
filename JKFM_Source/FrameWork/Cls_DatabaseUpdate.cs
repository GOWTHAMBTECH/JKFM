using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JKFM_Source
{
    class Cls_DatabaseUpdate
    {

        #region Declaration
        SAPbouiCOM.Form objform;
        ClsSBO objSBOAPI;
        #endregion        

        #region Constructor
        public Cls_DatabaseUpdate(ClsSBO objSBO)
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

                            if (pval.ItemUID == "3")
                            {                                
                                DialogResult Dialog;
                                Dialog = (DialogResult)objSBOAPI.SBO_Appln.MessageBox("Database will be updated. Do you wish to Proceed", 1, "Yes", "No");
                                if (Dialog == DialogResult.OK)
                                {
                                    objSBOAPI.objMain.CreateTables();
                                    objSBOAPI.SBO_Appln.StatusBar.SetText("Database will be updated successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                    try
                                    {
                                        objform.Close();
                                    }
                                    catch (Exception)
                                    {
                                    }
                                }
                            }
                            else if (pval.ItemUID == "4")
                            {
                                SAPbobsCOM.Recordset Orec;
                                Orec = (SAPbobsCOM.Recordset)objSBOAPI.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                string Str = "";
                                if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                {
                                    Str = "SELECT * FROM OUTB WHERE \"TableName\" = 'AV_QRBPT'";
                                }
                                else
                                {
                                    Str = "SELECT * FROM OUTB WHERE TableName = 'AV_QRBPT'";
                                }
                                 
                                Orec.DoQuery(Str);
                                if (Orec.RecordCount > 0)
                                {
                                    if (objSBOAPI.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                                    {
                                        objSBOAPI.create_queryCat("AddonQuery_Hana", "YYYYYYYYYYYYYYYYYYYY");
                                        Str = "SELECT CAST(ROW_NUMBER() OVER ( ORDER BY A.\"QName\") AS varchar) AS \"#\", '' AS \"Select\", A.\"QName\" AS \"QueryName\" FROM OUQR A INNER JOIN OQCN B ON A.\"QCategory\" = B.\"CategoryId\" WHERE B.\"CatName\" = 'AddonQuery_Hana'";
                                    }
                                    else
                                    {
                                        objSBOAPI.create_queryCat("AddonQuery_SQL", "YYYYYYYYYYYYYYYYYYYY");
                                        Str = "SELECT CONVERT(VARCHAR,ROW_NUMBER()OVER(ORDER BY A.QName),5) AS '#',''[Select],A.QName'QueryName' FROM OUQR A INNER JOIN OQCN B ON A.QCategory = B.CategoryId WHERE B.CatName = 'AddonQuery_SQL'";
                                    }
                                     
                                    Orec.DoQuery(Str);
                                    if (Orec.RecordCount > 0)
                                    {
                                        SAPbouiCOM.Form oForm = objSBOAPI.LoadForm("QueryManagerUpdate.xml", "AV_QRUPF", true);
                                        SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("Item_0").Specific;
                                        objform.Freeze(true);
                                        oGrid.DataTable.ExecuteQuery(Str);
                                        if (oGrid.DataTable.IsEmpty == false)
                                        {
                                            oGrid.Columns.Item("RowsHeader").Width = 0;
                                            oGrid.Columns.Item("#").Editable = false;
                                            oGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                                            oGrid.Columns.Item("QueryName").Editable = false;
                                        }
                                        oGrid.AutoResizeColumns();
                                        objform.Freeze(false);
                                    }
                                    else
                                    {
                                        objSBOAPI.objMain.CreateQueries();
                                    }
                                }
                                else
                                {
                                    objSBOAPI.SBO_Appln.MessageBox("Please do Database Update before QueryManager Update");
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
                }
                else if (pval.BeforeAction == false)
                {
                    if (pval.MenuUID == "AV_DBUPM")
                    {
                        objform = objSBOAPI.LoadForm("DatabaseUpdate.xml", "AV_DBUPF");
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
