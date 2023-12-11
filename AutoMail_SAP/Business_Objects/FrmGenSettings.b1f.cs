using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AutoMail_SAP.Common;
using SAPbouiCOM.Framework;

namespace AutoMail_SAP.Business_Objects
{
    [FormAttribute("138", "Business_Objects/FrmGenSettings.b1f")]
    class FrmGenSettings : SystemFormBase
    {
        public FrmGenSettings()
        {
        }
        public static SAPbouiCOM.Form oForm;
        private string strSQL;
        private SAPbobsCOM.Recordset objRs;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("fldrmail").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("lsmtpser").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("tsmtpser").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("lport").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("tport").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("logauth").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("lusernam").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("tusernam").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("lpass").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("tpass").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("lsapserv").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("tsapserv").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("mtxcont").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("ldbuser").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("tdbuser").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("ldbpass").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("tdbpass").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Folder Folder0;

        private void OnCustomInitialize()
        {
            try
            {
                oForm = clsModule.objaddon.objapplication.Forms.GetForm("138", 1);
                oForm = clsModule.objaddon.objapplication.Forms.ActiveForm;
                oForm.PaneLevel = 26;
                Folder0.GroupWith("2009");
                clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "sub", "#");

                Folder0.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                Matrix0.Columns.Item("trantype").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            }
            catch (Exception)
            {
                throw;
            }
        }

        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.StaticText StaticText7;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.Button Button0;

        private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
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
                //string live;
                SAPbobsCOM.GeneralService oGeneralService;
                SAPbobsCOM.GeneralData oGeneralData;
                SAPbobsCOM.GeneralDataParams oGeneralParams;
                SAPbobsCOM.GeneralDataCollection oGeneralDataCollection;
                SAPbobsCOM.GeneralData oChild;

                oGeneralService = clsModule.objaddon.objcompany.GetCompanyService().GetGeneralService("MAILCFG");
                oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralDataCollection = oGeneralData.Child("ATAMCFG1");
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
                oGeneralData.SetProperty("U_SMTPServ", oForm.DataSources.UserDataSources.Item("smptser").Value);
                oGeneralData.SetProperty("U_SMTPPort", oForm.DataSources.UserDataSources.Item("smptport").Value);
                oGeneralData.SetProperty("U_MailUser", oForm.DataSources.UserDataSources.Item("username").Value);
                oGeneralData.SetProperty("U_MailPass", oForm.DataSources.UserDataSources.Item("pass").Value);
                oGeneralData.SetProperty("U_SAPServ", oForm.DataSources.UserDataSources.Item("sapserver").Value);
                oGeneralData.SetProperty("U_DBUser", oForm.DataSources.UserDataSources.Item("dbuser").Value);
                oGeneralData.SetProperty("U_DBPass", oForm.DataSources.UserDataSources.Item("dbpass").Value);

                oChild = oGeneralDataCollection.Add();

                for (int i = 1; i <= Matrix0.VisualRowCount; i++)
                {
                    if (((SAPbouiCOM.EditText)Matrix0.Columns.Item("sub").Cells.Item(i).Specific).String != "")
                    {
                        if (i > oGeneralData.Child("ATEICFG1").Count)
                        {
                            oGeneralData.Child("ATEICFG1").Add();
                        }

                        oGeneralData.Child("ATEICFG1").Item(i - 1).SetProperty("U_TranType", ((SAPbouiCOM.ComboBox)Matrix0.Columns.Item("trantype").Cells.Item(i).Specific).Selected.Value);
                        oGeneralData.Child("ATEICFG1").Item(i - 1).SetProperty("U_Path", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("path").Cells.Item(i).Specific).String);
                        oGeneralData.Child("ATEICFG1").Item(i - 1).SetProperty("U_Param", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("param").Cells.Item(i).Specific).String);
                        oGeneralData.Child("ATEICFG1").Item(i - 1).SetProperty("U_ParamVal", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("paramval").Cells.Item(i).Specific).String);
                        oGeneralData.Child("ATEICFG1").Item(i - 1).SetProperty("U_MailSub", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("sub").Cells.Item(i).Specific).String);
                        oGeneralData.Child("ATEICFG1").Item(i - 1).SetProperty("U_MailCont", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("cont").Cells.Item(i).Specific).String);


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

    }
}
