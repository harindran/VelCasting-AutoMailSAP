using AutoMail_SAP.Common;
using CrystalDecisions.CrystalReports.Engine;
using System.Net.Mail;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CrystalDecisions.Shared;
using System.Data;

namespace AutoMail_SAP.Business_Objects
{
    class ClsARInvoice : clsAddon
    {
        public const string Formtype = "133";
        private SAPbouiCOM.Form oForm;
        private SAPbobsCOM.Recordset objRs;
        private SAPbouiCOM.DBDataSource dbDataSource_Head;
        string Query;

        #region ITEM EVENT
        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                //oForm = clsModule.objaddon.objapplication.Forms.Item(oFormUID);
                if (pVal.BeforeAction)
                {
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                            break;                        
                    }
                }
                else
                {
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:                            
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
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            if (BusinessObjectInfo.ActionSuccess == false) return;                           
                            
                            //Auto_Mail("","OINV", "13");
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:  // Velcasting said to trigger mail in Add Mode
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            if (BusinessObjectInfo.ActionSuccess== false) return;
                            string strsql;
                            dbDataSource_Head = oForm.DataSources.DBDataSources.Item("OINV");
                            strsql = clsModule.objaddon.objglobalmethods.getSingleValue("select 1 from OINV where \"DocEntry\"=" + dbDataSource_Head.GetValue("DocEntry", 0) + " and \"CANCELED\" ='Y'");
                            if (strsql == "1") return;
                            
                            //Query = "Select (Select \"BPLName\" from OBPL where \"BPLId\"= T0.\"BPLId\") \"Branch\",T0.\"DocNum\" \"Invoice No\",T0.\"DocDate\" \"Invoice Date\",";
                            //Query += "\n T0.\"CardName\",T0.\"NumAtCard\" \"Customer PO No\",";
                            //Query += "\n Case when T0.\"DocType\"='I' then T1.\"ItemCode\" Else (select \"ServCode\" from OSAC where \"AbsEntry\"=T1.\"SacEntry\") End \"Item Code\",";
                            //Query += "\n T1.\"Dscription\" \"Item Description\",T1.\"Quantity\" \"Quantity\",T1.\"Price\" \"Unit Price\",T1.\"TaxCode\" \"Tax Code\"";
                            //Query += "\n from OINV T0 join INV1 T1 on  T0.\"DocEntry\"=T1.\"DocEntry\" where T0.\"DocEntry\"="+ dbDataSource_Head.GetValue("DocEntry", 0) + ";";
                            //Query = Query.Insert(Query.Length, " where T0.\"DocEntry\"=" + dbDataSource_Head.GetValue("DocEntry", 0) + "");
                            Auto_Mail(BusinessObjectInfo.FormUID, dbDataSource_Head.GetValue("DocEntry", 0), "OINV","13");
                            //if (Auto_Mail(dbDataSource_Head.GetValue("DocEntry", 0),"13") == true)
                            //{
                            //    clsModule.objaddon.objapplication.StatusBar.SetText("Mail Sent Successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                            //}
                            //else
                            //{ clsModule.objaddon.objapplication.StatusBar.SetText("Failed to Sending Mail...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            //}
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

        #region FUNCTIONS
        
        public bool Auto_Mail(string FormUID,string DocEntry, string HeaderTable, string ObjType)
        {
            try
            {
                oForm = clsModule.objaddon.objapplication.Forms.Item(FormUID);
                string Title = oForm.Title.ToUpper();
                if (Title.EndsWith("CANCELLATION") == true) return true;
                ReportDocument cryRpt = new ReportDocument();
                string FromMail_id , FromMail_Password, Mail_Host, Mail_Port, strquery,tranquery;
                bool flag = false;
                string strsql;
                SAPbobsCOM.Recordset objInvoicerec, objcc;
                SAPbobsCOM.Recordset objrsupdate;
                string Mailbody, ServerName, CompanyDb, DBUserName, DbPassword;
                if (DocEntry != "") {
                    strquery = clsModule.objaddon.objglobalmethods.getSingleValue("Select 1 from " + HeaderTable + " T0 where T0.\"U_AutoMail\"='Y' and T0.\"DocEntry\"=" + DocEntry + " ");
                    if (strquery == "1") return true;
                }
                //clsModule.objaddon.objglobalmethods.WriteErrorLog("Started...");
                strquery = "Select T0.\"U_SMTPServ\",T0.\"U_SMTPPort\",T0.\"U_MailUser\",T0.\"U_MailPass\",T0.\"U_SAPServ\",T0.\"U_DBUser\",T0.\"U_DBPass\",";
                strquery += "\n T1.\"U_TranType\",T1.\"U_Path\",T1.\"U_Param\",T1.\"U_ParamVal\",T1.\"U_MailSub\",T1.\"U_MailCont\",T1.\"U_Query\"";
                strquery += "\n from \"@ATAMCFG\" T0 join \"@ATAMCFG1\" T1 on T0.\"Code\"=T1.\"Code\" where T0.\"Code\"='01' and T1.\"U_TranType\"='"+ ObjType + "'";
                //clsModule.objaddon.objglobalmethods.WriteErrorLog("First Query..."+ strquery);
                objRs = (SAPbobsCOM.Recordset)clsModule.objaddon. objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objcc = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objRs.DoQuery(strquery);
                //clsModule.objaddon.objglobalmethods.WriteErrorLog("First Query Executed..." );
                var Payroll_Report_FileName = Convert.ToString(objRs.Fields.Item("U_Path").Value);//System.Windows.Forms.Application.StartupPath + @"\" + "PaySlip_YMH.rpt"; 
                //Payroll_Report_FileName = @"E:\Chitra\Common Payroll\Dec 16\HRMS_Posting\PaySlip_YMH.rpt";//E:\Chitra\Seyoon_VelCasting\Auto Mail\layout1.rpt
                //Payroll_Report_FileName = @"\\newton.tmicloud.net\DB1SHARE\OEC_TEST\Attachments\CheckPrint\RptFile\CheckPrinting.rpt";
                ServerName = clsModule. objaddon.objcompany.Server;// objRs.Fields.Item("U_SAPServ").Value.ToString();// "WATSON.TMICLOUD.NET:30015"; 
                CompanyDb =clsModule.objaddon.objcompany.CompanyDB;// "OEC_TEST"; 
                DBUserName = objRs.Fields.Item("U_DBUser").Value.ToString();// "OECDBBR"; 
                DbPassword = objRs.Fields.Item("U_DBPass").Value.ToString();//"India@1947";
                //HTML Content
                //<!DOCTYPE html><html><body><p>Dear Sir/Madam,</p> <p> <br> Purpose of this Debit note/Invoice is generated for Shortage/Rejection basis. In case of any query, Please inform us within 24 Hrs (Due To EInvoice). </br></p><p><br>This is an auto generated email. Please do not reply to this email. Thank you! </br></p></body></html>
                if (objRs.RecordCount == 0) {
                    clsModule.objaddon.objapplication.MessageBox("Auto Mail Configuration Data Not Found...", 0, "OK");
                    //clsModule.objaddon.objapplication.StatusBar.SetText("Auto Mail Configuration Data Not Found...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                FromMail_id = objRs.Fields.Item("U_MailUser").Value.ToString(); // "saptech18@mukeshinfoserve.com"
                FromMail_Password = objRs.Fields.Item("U_MailPass").Value.ToString(); // "D@rloo@30895"
                Mail_Host = objRs.Fields.Item("U_SMTPServ").Value.ToString(); // "smtp-mail.outlook.com"
                Mail_Port = objRs.Fields.Item("U_SMTPPort").Value.ToString(); // "587"
                if (FromMail_id == "" | FromMail_Password == "" | Mail_Host == "" | Mail_Port == "")
                    return false;
                if (Payroll_Report_FileName != "")
                {                   
                    cryRpt.Load(Payroll_Report_FileName);
                    cryRpt.DataSourceConnections[0].SetConnection(ServerName, CompanyDb, false);
                    cryRpt.DataSourceConnections[0].SetLogon(DBUserName, DbPassword);
                    try
                    {
                        cryRpt.Refresh();
                        cryRpt.VerifyDatabase();
                    }
                    catch (Exception ex)
                    {
                        clsModule.objaddon.objapplication.StatusBar.SetText("Verify Database: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                }
               
                var Email = new System.Net.Mail.MailMessage();
                var MailServer = new System.Net.Mail.SmtpClient();
                
                strquery = "Select T0.\"U_AutoMail\",T0.\"BPLId\",";
                if (ObjType == "13")
                    strquery += "\n (Select \"U_EmailTo\" from CRD1 where \"CardCode\"=T0.\"CardCode\" and \"Address\"=T0.\"ShipToCode\" and \"AdresType\"='S' and \"U_EmailTo\" is not null) \"ToEmail\",(Select \"U_Emailcc\" from CRD1 where \"CardCode\"=T0.\"CardCode\" and \"Address\"=T0.\"ShipToCode\" and \"AdresType\"='S' and \"U_Emailcc\" is not null) \"CcEmail\",";
                else if (ObjType == "46")
                    strquery += "\n (Select \"E_Mail\" from OCRD where \"CardCode\"=T0.\"CardCode\" and \"E_Mail\" is not null) \"ToEmail\",'' \"CcEmail\", ";
                else
                    strquery += "\n (Select \"U_EmailTo\" from CRD1 where \"CardCode\"=T0.\"CardCode\" and \"Address\"=T0.\"PayToCode\" and \"AdresType\"='B' and \"U_EmailTo\" is not null) \"ToEmail\",(Select \"U_Emailcc\" from CRD1 where \"CardCode\"=T0.\"CardCode\" and \"Address\"=T0.\"PayToCode\" and \"AdresType\"='B' and \"U_Emailcc\" is not null) \"CcEmail\",";
                strquery += "\n (Select \"E_Mail\" from OUSR where \"USER_CODE\"='" + Convert.ToString(clsModule.objaddon.objcompany.UserName) + "' and \"E_Mail\" is not null) \"User E_Mail\",";
                strquery += "\n (Select \"U_NAME\" from OUSR where \"USER_CODE\"='" + Convert.ToString(clsModule.objaddon.objcompany.UserName) + "' and \"E_Mail\" is not null) \"User Name\",";
                strquery += "\n T0.\"DocNum\",T0.\"DocEntry\",T0.\"CardCode\",T0.\"CardName\"";
                strquery += "\n from "+ HeaderTable + " T0 where T0.\"U_AutoMail\"='N' ";
                if (DocEntry != "") strquery += "\n and T0.\"DocEntry\"=" + DocEntry + "";
                if (DocEntry == "") strquery += "\n and T0.\"DocDate\">='20221201'"; //Nov 29th 2022
                strquery += "\n Order By T0.\"DocEntry\" desc";
                //clsModule.objaddon.objglobalmethods.WriteErrorLog("Second Query..."+ strquery);
                objInvoicerec = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objInvoicerec.DoQuery(strquery);
                
                if (objInvoicerec.RecordCount == 0) {
                    //clsModule.objaddon.objapplication.MessageBox("No Data Found for Sending Mail...", 0, "OK");
                    //clsModule.objaddon.objapplication.StatusBar.SetText("No Data Found...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                //clsModule.objaddon.objglobalmethods.WriteErrorLog("Second Query Executed...");
                if (DocEntry != "")
                if (Convert.ToString(objInvoicerec.Fields.Item("ToEmail").Value) == "")
                {
                    clsModule.objaddon.objapplication.MessageBox("To Mail ID Not Found... " + Convert.ToString(objInvoicerec.Fields.Item("CardName").Value), 0, "OK");
                    //clsModule.objaddon.objapplication.StatusBar.SetText("Customer Mail ID Not Found...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else if (Convert.ToString(objInvoicerec.Fields.Item("User E_Mail").Value) == "")
                {
                    clsModule.objaddon.objapplication.MessageBox("User Mail ID Not Found..." + Convert.ToString(objInvoicerec.Fields.Item("User Name").Value), 0, "OK");
                    //clsModule.objaddon.objapplication.StatusBar.SetText("User Mail ID Not Found...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                clsModule.objaddon.objapplication.StatusBar.SetText("Mail Sending. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                for (int i = 0; i <objInvoicerec.RecordCount ; i++)
                    {
                    if (Convert.ToString(objInvoicerec.Fields.Item("ToEmail").Value) == ""| Convert.ToString(objInvoicerec.Fields.Item("User E_Mail").Value) == "")
                        continue;                
                    try
                    {
                        MailServer.Host = Mail_Host;
                        MailServer.Port = Convert.ToInt32(Mail_Port);
                        MailServer.Credentials = new System.Net.NetworkCredential(FromMail_id, FromMail_Password);
                        MailServer.EnableSsl = true;
                        Email.From = new System.Net.Mail.MailAddress(Convert.ToString(objInvoicerec.Fields.Item("User E_Mail").Value));//FromMail_id

                        string[] Toemails = Convert.ToString(objInvoicerec.Fields.Item("ToEmail").Value).Split(Convert.ToChar(","));
                        string[] Ccemails = Convert.ToString(objInvoicerec.Fields.Item("CcEmail").Value).Split(Convert.ToChar(","));
                        foreach (string email in Toemails)
                        {
                            if (email !="")
                                Email.To.Add(new System.Net.Mail.MailAddress(email.Trim()));
                            //Email.To.Add(email.Trim());
                        }
                        foreach (string email in Ccemails)
                        {
                            if (email != "")
                                Email.CC.Add(new System.Net.Mail.MailAddress(email.Trim()));
                            //Email.CC.Add(email.Trim());
                        }
                        //Email.To.Add(new System.Net.Mail.MailAddress(Convert.ToString(objInvoicerec.Fields.Item("ToEmail").Value)));
                        strquery = "Select \"U_EMailID\" \"CCMailID\" from \"@MAILCC\" where \"U_TranType\"='" + ObjType + "' and \"U_Branch\"='" + objInvoicerec.Fields.Item("BPLId").Value + "' and \"U_BPCode\"='" + objInvoicerec.Fields.Item("CardCode").Value + "' and \"U_EMailID\" is not null ";
                        //strquery += "\n Union All";
                        //strquery += "\n Select \"E_MailL\" from OCPR where \"CardCode\"='"+ Convert.ToString(objInvoicerec.Fields.Item("CardCode").Value) + "' and \"E_MailL\"<>''";
                        objcc.DoQuery(strquery);
                        string[] Cc1emails = Convert.ToString(objcc.Fields.Item("CCMailID").Value).Split(Convert.ToChar(","));
                        if (objcc.RecordCount>0)
                            foreach (string email in Cc1emails)
                            {
                                if (email != "")
                                    Email.CC.Add(new System.Net.Mail.MailAddress(email.Trim()));
                            }
                        //for (int Rec = 0; Rec < objcc.RecordCount; Rec++) {
                        //    Email.CC.Add(new System.Net.Mail.MailAddress(Convert.ToString(objcc.Fields.Item("CCMailID").Value)));
                        //    objcc.MoveNext();
                        //}
                            
                        Email.Subject = objRs.Fields.Item("U_MailSub").Value.ToString();// "Invoice - " + objInvoicerec.Fields.Item("CardName").Value ;
                        Mailbody = objRs.Fields.Item("U_MailCont").Value.ToString();
                        tranquery = objRs.Fields.Item("U_Query").Value.ToString();
                        if (tranquery!="")
                        Mailbody = Mailbody.Replace("{Table Data}", ConvertQueryToHTML(tranquery.Insert(tranquery.Length, " where T0.\"DocEntry\"=" + DocEntry + ""),ObjType));
                        //Mailbody = Mailbody.Replace("{FromName}", Convert.ToString(objInvoicerec.Fields.Item("User Name").Value));
                        ////Mailbody = "Dear " + objInvoicerec.Fields.Item("CardName").Value + ",";
                        ////Mailbody += " ";
                        ////Mailbody += " test";
                        ////Mailbody +=  " ";
                        ////Mailbody += "This is Auto generated E-Mail from SAP Business One . Please do not reply to this message. Thank you! ";

                        Email.Body = Mailbody;
                        Email.IsBodyHtml = true;
                        Email.Priority = System.Net.Mail.MailPriority.Normal;
                        //string param = Convert.ToString(objRs.Fields.Item("U_Param").Value);// "@DocKey";
                        //cryRpt.SetParameterValue("DocKey@", Convert.ToInt32(objrs.Fields.Item("Year").Value)); // 2022 "@DocKey" Convert.ToString(objRs.Fields.Item("U_Param").Value.ToString())

                        if (Payroll_Report_FileName != "")
                        {
                            cryRpt.SetParameterValue(Convert.ToString(objRs.Fields.Item("U_Param").Value), objInvoicerec.Fields.Item("DocEntry").Value); // Convert.ToString(objRs.Fields.Item("U_Param").Value)                                                                                                                                                                                           // cryRpt.SetParameterValue("Emp@select Distinct T1.""U_empID"",T1.""U_empName"" from ""@MIPL_PPI1"" T1 where ifnull(T1.""U_empID"",'')<>''", 
                            Email.Attachments.Add(new Attachment(cryRpt.ExportToStream(ExportFormatType.PortableDocFormat), "DOC- " + objInvoicerec.Fields.Item("DocNum").Value + ".PDF"));
                        }
                           

                        MailServer.Send(Email);                        
                        strsql = "Update " + HeaderTable + " set \"U_AutoMail\"='Y',\"U_MailRem\"='Mail Sent!' where \"DocEntry\"='" + objInvoicerec.Fields.Item("DocEntry").Value + "' and \"U_AutoMail\"='N' ";
                        objrsupdate = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        objrsupdate.DoQuery(strsql);
                        //Update_Invoice(Convert.ToString(objInvoicerec.Fields.Item("DocEntry").Value),SAPbobsCOM.BoObjectTypes.oInvoices);
                        flag = true;
                    }
                    catch (Exception ex)
                    {                        
                        flag = false;
                        strsql = "Update " + HeaderTable + " set \"U_AutoMail\"='N',\"U_MailRem\"='Failed-' || '"+ ex.Message + "' where \"DocEntry\"='" + objInvoicerec.Fields.Item("DocEntry").Value + "' ";
                        objrsupdate = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        objrsupdate.DoQuery(strsql);
                        clsModule.objaddon.objapplication.StatusBar.SetText("Auto_Mail Transaction: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                    finally
                    {
                        if (Email !=null)
                        Email.Dispose();
                        MailServer = null;
                     }
                    objInvoicerec.MoveNext();
                }
                if (flag == true)
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Mail Sent Successfully!...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    return true;
                }
                else
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Failed to Sending Mail...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    return false;
                }            

            }
            catch (Exception ex)
            {                
                clsModule.objaddon.objapplication.StatusBar.SetText("Auto_Mail" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
        }
        
        private string Getvalue_webconfig(string key)
        {
            try
            {
                string strConnectionString = System.Configuration.ConfigurationManager.AppSettings.Get(key);
                return strConnectionString;
            }
            catch (Exception ex)
            {
                // Interaction.MsgBox(ex.ToString());
                return "";
            }
        }

        public bool Update_Invoice(string DocEntry, SAPbobsCOM.BoObjectTypes trantype) {
            try
            {
                SAPbobsCOM.Documents objSalesInvoice;
                int Retval;
                objSalesInvoice = (SAPbobsCOM.Documents)clsModule.objaddon.objcompany.GetBusinessObject(trantype);//SAPbobsCOM.BoObjectTypes.oInvoices
                objSalesInvoice.GetByKey(Convert.ToInt32(DocEntry));
                if (Convert.ToString(objSalesInvoice.UserFields.Fields.Item("U_AutoMail").Value) == "N")
                {
                    objSalesInvoice.UserFields.Fields.Item("U_AutoMail").Value = "Y";
                    //objSalesInvoice.UserFields.Fields.Item("U_MailRem").Value = "Mail Sent!";
                }

                Retval = objSalesInvoice.Update();
                if (Retval == 0)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objSalesInvoice);
                }
                else
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Invoice: " + clsModule.objaddon.objcompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }
            catch (Exception)
            {

                //throw;
            }           

            return true;
        }

        public  string ConvertQueryToHTML(string Query,string ObjType)
        {
            string html = "<table border = '1'>";
            double doctotal=0,advance=0;
            SAPbobsCOM.Recordset recordset;
            recordset = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            recordset.DoQuery(Query);

            if (recordset.RecordCount == 0) return "";                      
            //add header row
            html += "<tr>";
            for (int i = 0; i < recordset.Fields.Count; i++)
            {
                if (recordset.Fields.Item(i).Name == "DocTotal"||recordset.Fields.Item(i).Name == "Advance") continue;
                html += "<th>" + recordset.Fields.Item(i).Name + "</th>";
            }
               
            html += "</tr>";
            //add rows
            for (int i = 0; i < recordset.RecordCount; i++)
            {
                html += "<tr>";
                for (int j = 0; j < recordset.Fields.Count; j++)
                {
                    if (recordset.Fields.Item(j).Name == "DocTotal" || recordset.Fields.Item(j).Name == "Advance") continue;
                    html += "<td>" + recordset.Fields.Item(j).Value.ToString() + "</td>";
                    if (doctotal == 0) doctotal = Convert.ToDouble(recordset.Fields.Item("DocTotal").Value);
                    if (ObjType=="46") if (advance == 0) advance = Convert.ToDouble(recordset.Fields.Item("Advance").Value);
                }             
                
                recordset.MoveNext();
                html += "</tr>";                
            }
            if (ObjType == "13")
                html += "<tfoot><tr><td colspan='13' style='text-align:right'>Total</td><td> "+ doctotal + " </td></tr></tfoot>";
            if (ObjType == "20"| ObjType == "22" | ObjType == "19")
                html += "<tfoot><tr><td colspan='12' style='text-align:right'>Total</td><td> " + doctotal + " </td></tr></tfoot>";
            if (ObjType == "46")
                html += "<tfoot><tr><td colspan='5' style='text-align:right'>Advance</td><td> " + advance + " </td></tr><tr><td colspan='5' style='text-align:right'>Total</td><td> " + doctotal + " </td></tr></tfoot>";
            html += "</table>";
            return html;
        }

        #endregion
    }
}
