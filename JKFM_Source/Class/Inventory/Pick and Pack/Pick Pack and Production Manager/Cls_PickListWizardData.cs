using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Media;

namespace JKFM_Source
{
    class Cls_PickListWizardData
    {
        #region Declaration
        SAPbouiCOM.Form objform;
        ClsSBO objSBOAPI;
        SAPbouiCOM.DBDataSource oDBDSHeaderPickWizard;
        #endregion        

        #region Constructor
        public Cls_PickListWizardData(ClsSBO objSBO)
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
                        case SAPbouiCOM.BoEventTypes.et_CLICK:
                            switch(pval.ItemUID)
                            {
                                case "26":
                                    string Unit1 = Convert.ToString(objform.Items.Item("234000035").Specific.Value);
                                    string Unit2 = Convert.ToString(objform.Items.Item("234000066").Specific.Value);
                                    if (Unit1 == string.Empty || Unit2 == string.Empty)
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("Unit Shouldn't be empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                        SystemSounds.Asterisk.Play();
                                        bubbleevent = false;
                                    }
                                    else if (Unit1 != Unit2)
                                    {
                                        objSBOAPI.SBO_Appln.StatusBar.SetText("Single Unit Should Proceed", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                        SystemSounds.Asterisk.Play();
                                        bubbleevent = false;
                                    }
                                    break;
                            }
                            break;

                        //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        //    objform = objSBOAPI.SBO_Appln.Forms.GetForm("80", -1);
                        //    oDBDSHeaderPickWizard = objform.DataSources.DBDataSources.Item(0);
                        //    break;
                    }
                }
                else
                {
                    switch (pval.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                            objform = objSBOAPI.SBO_Appln.Forms.GetForm("80", -1);
                            oDBDSHeaderPickWizard = objform.DataSources.DBDataSources.Item(0);
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
    }
}
