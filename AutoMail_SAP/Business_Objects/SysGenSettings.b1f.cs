using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AutoMail_SAP.Common;
using SAPbouiCOM.Framework;

namespace AutoMail_SAP.Business_Objects
{
    [FormAttribute("138", "Business_Objects/SysGenSettings.b1f")]
    class SysGenSettings : SystemFormBase
    {
        public static SAPbouiCOM.Form oForm;
        private string strSQL;
        private SAPbobsCOM.Recordset objRs;
        public SysGenSettings()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("fldrmail").Specific));
            this.Folder0.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder0_PressedAfter);
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("lsmtpser").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("tsmtpser").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("lsmtport").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("tsmtport").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("lmailid").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("tmailid").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("lmailpas").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("tmailpas").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("lsapser").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("tsapser").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("ldbuser").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("tdbuser").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("ldbpass").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("tdbpass").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("mtxcont").Specific));
            this.Matrix0.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix0_ValidateAfter);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("btntmail").Specific));
            this.Button1.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button1_ClickBefore);
            this.Button1.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button1_ClickAfter);
            this.Folder1 = ((SAPbouiCOM.Folder)(this.GetItem("129").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);

        }


        private void OnCustomInitialize()
        {
            try
            {
                oForm = clsModule.objaddon.objapplication.Forms.GetForm("138", 1);
                //oForm = clsModule.objaddon.objapplication.Forms.ActiveForm;
                oForm.PaneLevel = 26;
                Folder0.GroupWith("2009");
                clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "sub", "#");
                Folder1.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                //Folder0.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                Matrix0.Columns.Item("trantype").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                Matrix0.AutoResizeColumns();
            }
            catch (Exception)
            {
                //throw;
            }
        }

        #region Fields
        
        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Folder Folder1;

        #endregion

        private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                oForm = clsModule.objaddon.objapplication.Forms.ActiveForm;
                if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) { return; }
                //objRs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                if (AutoMail_Config())
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Data Saved Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
            }
            catch (Exception)
            {

                //throw;
            }

        }

        private bool AutoMail_Config()
        {
            try
            {
                bool Flag = false;
                SAPbobsCOM.GeneralService oGeneralService;
                SAPbobsCOM.GeneralData oGeneralData;
                SAPbobsCOM.GeneralDataParams oGeneralParams;
                SAPbobsCOM.GeneralDataCollection oGeneralDataCollection;
                SAPbobsCOM.GeneralData oChild;

                oGeneralService = clsModule.objaddon.objcompany.GetCompanyService().GetGeneralService("MAILCFG");
                oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralDataCollection = oGeneralData.Child("ATAMCFG1");
                oForm = clsModule.objaddon.objapplication.Forms.ActiveForm;
                try
                {
                    oGeneralParams.SetProperty("Code", "01");
                    oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                    Flag = true;
                }
                catch (Exception ex)
                {
                    Flag = false;
                }

                oGeneralData.SetProperty("Code", "01");
                oGeneralData.SetProperty("Name", "01");
                oGeneralData.SetProperty("U_SMTPServ", EditText0.Value );//oForm.DataSources.UserDataSources.Item("smtpser").Value
                oGeneralData.SetProperty("U_SMTPPort", EditText1.Value );//oForm.DataSources.UserDataSources.Item("smptport").Value
                oGeneralData.SetProperty("U_MailUser", EditText2.Value );//oForm.DataSources.UserDataSources.Item("mailid").Value
                oGeneralData.SetProperty("U_MailPass", EditText3.Value );//oForm.DataSources.UserDataSources.Item("mailpass").Value
                oGeneralData.SetProperty("U_SAPServ", EditText4.Value );//oForm.DataSources.UserDataSources.Item("sapserv").Value
                oGeneralData.SetProperty("U_DBUser", EditText5.Value );//oForm.DataSources.UserDataSources.Item("dbuser").Value
                oGeneralData.SetProperty("U_DBPass", EditText6.Value );//oForm.DataSources.UserDataSources.Item("dbpass").Value

                oChild = oGeneralDataCollection.Add();

                for (int i = 1; i <= Matrix0.VisualRowCount; i++)
                {
                    if (((SAPbouiCOM.EditText)Matrix0.Columns.Item("sub").Cells.Item(i).Specific).String != "")
                    {
                        if (i > oGeneralData.Child("ATAMCFG1").Count)
                        {
                            oGeneralData.Child("ATAMCFG1").Add();
                        }

                        oGeneralData.Child("ATAMCFG1").Item(i - 1).SetProperty("U_TranType", ((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("trantype").Cells.Item(i).Specific).Selected.Value);
                        oGeneralData.Child("ATAMCFG1").Item(i - 1).SetProperty("U_Path", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("path").Cells.Item(i).Specific).String);
                        oGeneralData.Child("ATAMCFG1").Item(i - 1).SetProperty("U_Param", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("param").Cells.Item(i).Specific).String);
                        //oGeneralData.Child("ATAMCFG1").Item(i - 1).SetProperty("U_ParamVal", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("paramval").Cells.Item(i).Specific).String);
                        oGeneralData.Child("ATAMCFG1").Item(i - 1).SetProperty("U_MailCont", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("cont").Cells.Item(i).Specific).String);
                        oGeneralData.Child("ATAMCFG1").Item(i - 1).SetProperty("U_MailSub", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("sub").Cells.Item(i).Specific).String);
                        oGeneralData.Child("ATAMCFG1").Item(i - 1).SetProperty("U_Query", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("query").Cells.Item(i).Specific).String);

                    }
                }
                if (Flag == true)
                {
                    oGeneralService.Update(oGeneralData);
                    return true;
                }
                else
                {
                    oGeneralParams = oGeneralService.Add(oGeneralData);
                    return true;
                }
            }


            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
        }

        private void Matrix0_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                //if (pVal.InnerEvent == false) return;
                switch (pVal.ColUID)
                {
                    case "cont":
                        clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "cont", "#");
                        break;
                }

            }
            catch (Exception)
            {
                //throw;
            }

        }

        private void Folder0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                oForm = clsModule.objaddon.objapplication.Forms.ActiveForm;
                oForm.PaneLevel = 26;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

            }
            catch (Exception)
            {
                //throw;
            }

        }

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                oForm = clsModule.objaddon.objapplication.Forms.GetForm("138", pVal.FormTypeCount);
                objRs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                if (clsModule.objaddon.HANA == true)
                {
                    strSQL = "Select T0.\"U_SMTPServ\",T0.\"U_SMTPPort\",T0.\"U_MailUser\",T0.\"U_MailPass\",T0.\"U_SAPServ\",T0.\"U_DBUser\",T0.\"U_DBPass\",";
                    strSQL += "\n T1.\"U_TranType\",T1.\"LineId\",T1.\"U_Path\",T1.\"U_Param\",T1.\"U_ParamVal\",T1.\"U_MailSub\",T1.\"U_MailCont\",T1.\"U_Query\"";
                    strSQL += "\n from \"@ATAMCFG\" T0 join \"@ATAMCFG1\" T1 on T0.\"Code\"=T1.\"Code\" where T0.\"Code\"='01'";
                }

                else
                {
                    strSQL = "Select T0.U_SMTPServ,T0.U_SMTPPort,T0.U_MailUser,T0.U_MailPass,T0.U_SAPServ,T0.U_DBUser,T0.U_DBPass,";
                    strSQL += "\n T1.U_TranType,T1.LineId,T1.U_Path,T1.U_Param,T1.U_ParamVal,T1.U_MailSub,T1.U_MailCont,T1.U_Query";
                    strSQL += "\n from [@ATAMCFG] T0 join [@ATAMCFG1] T1 on T0.Code=T1.Code where T0.Code='01'";
                }
                objRs.DoQuery(strSQL);

                if (objRs.RecordCount > 0)
                {
                    EditText0.Value = objRs.Fields.Item("U_SMTPServ").Value.ToString();
                    EditText1.Value = objRs.Fields.Item("U_SMTPPort").Value.ToString();
                    EditText2.Value = objRs.Fields.Item("U_MailUser").Value.ToString();
                    EditText3.Value = objRs.Fields.Item("U_MailPass").Value.ToString();
                    EditText4.Value = objRs.Fields.Item("U_SAPServ").Value.ToString();
                    EditText5.Value = objRs.Fields.Item("U_DBUser").Value.ToString();
                    EditText6.Value = objRs.Fields.Item("U_DBPass").Value.ToString();

                    for (int i = 0; i < objRs.RecordCount; i++)
                    {
                        string dd = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("sub").Cells.Item(Matrix0.VisualRowCount).Specific).String;
                        if (Matrix0.VisualRowCount == 0 | (((SAPbouiCOM.EditText)Matrix0.Columns.Item("sub").Cells.Item(Matrix0.VisualRowCount).Specific).String != ""))
                            Matrix0.AddRow();
                        ((SAPbouiCOM.EditText)Matrix0.Columns.Item("#").Cells.Item(Matrix0.VisualRowCount).Specific).String = objRs.Fields.Item("LineId").Value.ToString();
                        ((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("trantype").Cells.Item(Matrix0.VisualRowCount).Specific).Select(objRs.Fields.Item("U_TranType").Value.ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                        ((SAPbouiCOM.EditText)Matrix0.Columns.Item("path").Cells.Item(Matrix0.VisualRowCount).Specific).String= objRs.Fields.Item("U_Path").Value.ToString();
                        ((SAPbouiCOM.EditText)Matrix0.Columns.Item("param").Cells.Item(Matrix0.VisualRowCount).Specific).String = objRs.Fields.Item("U_Param").Value.ToString();
                        //((SAPbouiCOM.EditText)Matrix0.Columns.Item("paramval").Cells.Item(Matrix0.VisualRowCount).Specific).String = objRs.Fields.Item("U_ParamVal").Value.ToString();
                        ((SAPbouiCOM.EditText)Matrix0.Columns.Item("sub").Cells.Item(Matrix0.VisualRowCount).Specific).String = objRs.Fields.Item("U_MailSub").Value.ToString();
                        ((SAPbouiCOM.EditText)Matrix0.Columns.Item("cont").Cells.Item(Matrix0.VisualRowCount).Specific).String = objRs.Fields.Item("U_MailCont").Value.ToString();
                        ((SAPbouiCOM.EditText)Matrix0.Columns.Item("query").Cells.Item(Matrix0.VisualRowCount).Specific).String = objRs.Fields.Item("U_Query").Value.ToString();
                        objRs.MoveNext();
                    }
                    Matrix0.AutoResizeColumns();
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

                }

            }
            catch (Exception)
            {

                throw;
            }

        }

        private bool Test_Mail()
        {
            try
            {                
                string FromMail_id = "";
                string FromMail_Password = "";
                string Mail_Host = "";
                string Mail_Port = "";     
               
                string Mailbody;               

                if (objRs.RecordCount == 0)
                    return false;
                FromMail_id = EditText2.Value;
                FromMail_Password = EditText3.Value;
                Mail_Host = EditText0.Value;
                Mail_Port = EditText1.Value;
                if (FromMail_id == "" | FromMail_Password == "" | Mail_Host == "" | Mail_Port == "")
                    return false;
               
                var Email = new System.Net.Mail.MailMessage();
                var MailServer = new System.Net.Mail.SmtpClient();               

                try
                {
                    MailServer.Host = Mail_Host;
                    MailServer.Port = Convert.ToInt32(Mail_Port);
                    MailServer.Credentials = new System.Net.NetworkCredential(FromMail_id, FromMail_Password);
                    MailServer.EnableSsl = true;
                    Email.From = new System.Net.Mail.MailAddress(FromMail_id);

                    Email.To.Add(new System.Net.Mail.MailAddress(Convert.ToString(FromMail_id)));
                    Email.Subject = "Test Mail" ;

                    Mailbody = "Hi Receiver,";                    
                    Mailbody += " \n\nTest mail";
                    Mailbody += " \n\n";
                    Mailbody += "This is Auto generated E-Mail from SAP Business One . Please do not reply to this message. Thank you! ";
                    Mailbody += " \n\n Regards,";
                    Mailbody += " \n Sender";
                    Email.Body = Mailbody;
                    //Email.IsBodyHtml = true;
                    Email.Priority = System.Net.Mail.MailPriority.Normal;                   
                    MailServer.Send(Email);
                  
                }
               
                catch (Exception ex)
                {
                    //throw;
                    clsModule.objaddon.objapplication.StatusBar.SetText("Test_Mail : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                finally
                {
                    if (Email != null)
                        Email.Dispose();
                    MailServer = null;
                }
                return true;
            }
            catch (Exception ex)
            {
                //throw;
                clsModule.objaddon.objapplication.StatusBar.SetText("Test_Auto_Mail" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
        }

        private SAPbouiCOM.Button Button1;

        private void Button1_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (Test_Mail() == true)
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Auto Test Mail Sent Successfully!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                else
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Error in test mail..." , SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Error" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Button1_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (EditText0.Value == "")
                {
                    clsModule.objaddon.objapplication.SetStatusBarMessage("SMTP Server is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    BubbleEvent = false;
                    return;
                }
                if (EditText1.Value == "")
                {
                    clsModule.objaddon.objapplication.SetStatusBarMessage("SMTP Port is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    BubbleEvent = false;
                    return;
                }
                if (EditText2.Value == "")
                {
                    clsModule.objaddon.objapplication.SetStatusBarMessage("Mail ID is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    BubbleEvent = false;
                    return;
                }
                if (EditText3.Value == "")
                {
                    clsModule.objaddon.objapplication.SetStatusBarMessage("Mail ID Password is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    BubbleEvent = false;
                    return;
                }
            }
            catch (Exception)
            {

                throw;
            }

        }

        

       
    }
}
