using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoMail_SAP.Common
{
    class clsMenuEvent
    {
        SAPbouiCOM.Form objform;
        //clsAddon objaddon;

        public void MenuEvent_For_StandardMenu(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                SAPbouiCOM.Form oUDFForm;
                if (pVal.MenuUID == "1287") {
                    objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                    oUDFForm = clsModule.objaddon.objapplication.Forms.Item(objform.UDFFormUID);
                    ((SAPbouiCOM.ComboBox)oUDFForm.Items.Item("U_AutoMail").Specific).Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    ((SAPbouiCOM.EditText)oUDFForm.Items.Item("U_MailRem").Specific).String = "";
                    objform.Items.Item("4").Click();
                }
                
                
                switch (clsModule. objaddon.objapplication.Forms.ActiveForm.TypeEx)
                {
                    //case "133":
                    case "-133"://AR Invoice
                    case "-142"://Purchase Order
                    case "-143": // Goods Receipt PO
                    case "-181": // AP Credit Memo
                    case "-426": // Outgoing Payment
                        {
                            // Default_Sample_MenuEvent(pVal, BubbleEvent)
                            if (pVal.BeforeAction == true)
                                return;
                            objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                            Default_Sample_MenuEvent(pVal, BubbleEvent);

                            break;
                        }
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void Default_Sample_MenuEvent(SAPbouiCOM.MenuEvent pval, bool BubbleEvent)
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                if (pval.BeforeAction == true)
                {
                }
                else
                {
                    SAPbouiCOM.Form oUDFForm;
                    try
                    {
                        oUDFForm = clsModule.objaddon.objapplication.Forms.Item(objform.UDFFormUID);
                    }
                    catch (Exception ex)
                    {
                        oUDFForm = objform;
                    }

                    switch (pval.MenuUID)
                    {
                        case "1281": // Find
                            {
                                oUDFForm.Items.Item("U_AutoMail").Enabled = true;
                                oUDFForm.Items.Item("U_MailRem").Enabled = true;
                                break;
                            }
                        case "1287":
                            {
                                if (oUDFForm.Items.Item("U_AutoMail").Enabled == false)
                                {
                                    oUDFForm.Items.Item("U_AutoMail").Enabled = true;
                                }
                                ((SAPbouiCOM.ComboBox)oUDFForm.Items.Item("U_AutoMail").Specific).Select("N",SAPbouiCOM.BoSearchKey.psk_ByValue);
                                ((SAPbouiCOM.EditText)oUDFForm.Items.Item("U_MailRem").Specific).String = "";
                                break;
                            }
                        default:
                            {
                                oUDFForm.Items.Item("U_AutoMail").Enabled = false;
                                oUDFForm.Items.Item("U_MailRem").Enabled = false;
                                break;
                            }
                    }
                }
            }
            catch (Exception ex)
            {
                // objaddon.objapplication.SetStatusBarMessage("Error in Standart Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            }
        }


    }
}
