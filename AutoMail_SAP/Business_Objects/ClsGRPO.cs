using AutoMail_SAP.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoMail_SAP.Business_Objects
{
    class ClsGRPO: clsAddon
    {
        public const string Formtype = "143";
        private SAPbouiCOM.Form oForm;
        private SAPbobsCOM.Recordset objRs;
        private SAPbouiCOM.DBDataSource dbDataSource_Head;
        string strsql;
        #region ITEM EVENT
        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm = clsModule.objaddon.objapplication.Forms.Item(oFormUID);
                if (pVal.BeforeAction)
                {
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DRAW:
                           
                            break;
                    }
                }
                else
                {
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DRAW:
                            
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:

                           
                            break;
                    }
                }
            }
            catch (Exception Ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                //BubbleEvent = false;
                return;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
        #endregion

        #region FORM DATA EVENT
        public override void FormData_Event(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                oForm = clsModule.objaddon.objapplication.Forms.Item(BusinessObjectInfo.FormUID);
                if (BusinessObjectInfo.BeforeAction)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            break;

                    }
                }
                else
                {
                    ClsARInvoice aRInvoice = new ClsARInvoice();
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            //if (BusinessObjectInfo.ActionSuccess == false) return;
                            //aRInvoice.Auto_Mail("", "OPDN", "20");
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            if (BusinessObjectInfo.ActionSuccess == false) return;
                            //if (((SAPbouiCOM.ComboBox)oForm.Items.Item("81").Specific).Selected.Value == "4") return;
                            dbDataSource_Head = oForm.DataSources.DBDataSources.Item("OPDN");
                            strsql = clsModule.objaddon.objglobalmethods.getSingleValue("select 1 from OPDN where \"DocEntry\"=" + dbDataSource_Head.GetValue("DocEntry", 0) + " and \"CANCELED\" ='Y'");
                            if (strsql == "1") return;                            
                            aRInvoice.Auto_Mail(BusinessObjectInfo.FormUID,dbDataSource_Head.GetValue("DocEntry", 0), "OPDN", "20");                            
                            break;

                    }
                }
            }
            catch (Exception Ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
                return;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
        #endregion

    }
}
