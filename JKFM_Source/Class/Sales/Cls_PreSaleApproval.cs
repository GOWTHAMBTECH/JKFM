using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM;

namespace JKFM_Source
{
    class Cls_PreSaleApproval
    {
        #region Declaration
        SAPbouiCOM.Form objform;
        ClsSBO objSBOAPI;

        int SubRowID_PlanningDetails;
        string SubRowID_DayID;
        SAPbouiCOM.DBDataSource oDBDSDetail;
        SAPbouiCOM.DBDataSource oDBDSHeader;
        SAPbouiCOM.Grid oGrid;
        SAPbouiCOM.DataTable dtloadgrid;
        bool boolFormLoaded;
        string sFormUID;
        string UDOID;

        #endregion        

        #region Constructor
        public Cls_PreSaleApproval(ClsSBO objSBO)
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
                    //    case SAPbouiCOM.BoEventTypes.et_CLICK:
                    //        switch(pval.ItemUID)
                    //        {
                    //            case "":
                    //                break;
                    //        }
                    //        break;
                    //}
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
                    switch(pval.MenuUID)
                    {
                        case "OPSA":
                            objform = objSBOAPI.LoadForm("PreSaleApproval.xml", "OPSA");
                            objform = objSBOAPI.SBO_Appln.Forms.Item("OPSA");
                            objform.Items.Item("Grid").Visible = true;
                            oGrid = (Grid)objform.Items.Item("Grid").Specific;
                            dtloadgrid = objform.DataSources.DataTables.Add("DataTable");
                            objform.Freeze(true);
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
    }
}
