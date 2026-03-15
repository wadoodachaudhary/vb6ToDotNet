using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

using System.Data.SqlClient;
using System.IO;

using System.Security.Permissions;


namespace AutoNotice
{
   
    public partial class AutoEventDetails : Form
    {
        public AutoEventDetails()
        {
            InitializeComponent();
        }

       
        /// <summary>
        /// Connection string
        /// </summary>
        private string sqlConnStr;

        /// <summary>
        /// Trigger Name
        /// </summary>
        private string strTriggerName;

        /// <summary>
        /// Table Name which Trigger is based on
        /// </summary>
        private string strTriggerTableName;

        /// <summary>
        /// Recipients
        /// </summary>
        private string strRecipient;
        private string strCC;

        /// <summary>
        /// Sender
        /// </summary>
        private string strFromAddress;

        /// <summary>
        /// Subject
        /// </summary>
        private string strSubject;

        /// <summary>
        /// Message
        /// </summary>
        private string strMessage;

        /// <summary>
        /// true if trigger is disabled
        /// </summary>
        private bool IsDisabled;
        
        /// <summary>
        /// 
        /// </summary>
        private bool IsBuyer;

        /// <summary>
        /// 
        /// </summary>
        private bool IsCoBuyer;

        /// <summary>
        /// 
        /// </summary>
        private bool IsSalesPerson;

        /// <summary>
        /// 
        /// </summary>
        private bool IsPM;

        /// <summary>
        /// 
        /// </summary>
        private bool IsSalesManager;

        /// <summary>
        /// 
        /// </summary>
        private bool IsEstimator;

        /// <summary>
        /// 
        /// </summary>
        private bool IsVendorPurchasing;

        /// <summary>
        /// 
        /// </summary>
        private bool IsVendorScheduling;

        /// <summary>
        /// 
        /// </summary>
        private bool IsVendorService;

        /// <summary>
        /// 
        /// </summary>
        private bool IsPurchaser;

        /// <summary>
        /// 
        /// </summary>
        private bool IsDCSalesPerson;

        /// <summary>
        /// 
        /// </summary>
        private bool IsAccountingAdmin;

        /// <summary>
        /// 
        /// </summary>
        private bool IsSystemAdmin;

        /// <summary>
        /// 
        /// </summary>
        private bool IsLender;

        /// <summary>
        /// 
        /// </summary>
        private bool IsTitleCompany;

        /// <summary>
        /// 
        /// </summary>
        private bool IsRealtor;

        /// <summary>
        /// 
        /// </summary>
        private bool IsBroker;

        /// <summary>
        /// 
        /// </summary>
        private bool IsSellingRealtor;
        
        /// <summary>
        /// Attachment
        /// </summary>
        private string strAttachment;

        /// <summary>
        /// Report
        /// </summary>
        private string strReport;

        /// <summary>
        /// Flag, if Cancel button has been clicked or not
        /// </summary>
        private bool IsCanceled;

        /// <summary>
        /// Flag, if data have been modified
        /// </summary>
        private bool IsModified;

        /// <summary>
        /// Array of report parameters
        /// </summary>
        ArrayList reportParams = new ArrayList();

        /// <summary>
        /// Flag showing if form is being loaded or not.
        /// </summary>
        private bool IsLoading;

        /// <summary>
        /// number of renamed parameters
        /// </summary>
        private int renamedParams;

        /// <summary>
        /// The ODBC DSN name
        /// </summary>
        private string DataSourceName;

        /// <summary>
        /// Database user name
        /// </summary>
        private string UID;

        /// <summary>
        /// Database password
        /// </summary>
        private string PWD;


        /// <summary>
        /// Sets DSN name
        /// </summary>
        public string DSN_NAME
        {
            set
            {
                DataSourceName = value;
            }
        }

        /// <summary>
        /// Sets database user name
        /// </summary>
        public string DB_USER
        {
            set
            {
                UID = value;
            }
        }

        /// <summary>
        /// Sets database password
        /// </summary>
        public string DB_PASSWORD
        {
            set
            {
                PWD = value;
            }
        }

        /// <summary>
        /// Sets/Gets connection string
        /// </summary>
        public string ConnectionString
        {
            get
            {
                return sqlConnStr;
            }
            set
            {
                sqlConnStr = value;
            }
        }

        /// <summary>
        /// true if Cancel button has been clicked
        /// </summary>
        public bool Canceled
        {
            get
            {
                return IsCanceled;
            }
        }

        /// <summary>
        /// true if data have been modified
        /// </summary>
        public bool Modified
        {
            get
            {
                return IsModified;
            }
        }

        /// <summary>
        /// Sets/Gets report parameters
        /// </summary>
        public ArrayList ReportParameters
        {
            get
            {
                return reportParams;
            }
            set
            {
                reportParams = value;
            }
        }

        /// <summary>
        /// true if trigger is disabled
        /// </summary>
        public bool Disabled
        {
            get
            {
                return IsDisabled;
            }
            set
            {
                IsDisabled = value;
            }
        }

        /// <summary>
        /// true if Buyer recieves a notice
        /// </summary>
        public bool Buyer
        {
            get
            {
                return IsBuyer;
            }
            set
            {
                IsBuyer = value;
            }
        }

        /// <summary>
        /// true if CoBuyer recieves a notice
        /// </summary>
        public bool CoBuyer
        {
            get
            {
                return IsCoBuyer;
            }
            set
            {
                IsCoBuyer = value;
            }
        }

        /// <summary>
        /// true if SalesPerson recieves a notice
        /// </summary>
        public bool SalesPerson
        {
            get
            {
                return IsSalesPerson;
            }
            set
            {
                IsSalesPerson = value;
            }
        }

        /// <summary>
        /// true if PM recieves a notice
        /// </summary>
        public bool PM
        {
            get
            {
                return IsPM;
            }
            set
            {
                IsPM = value;
            }
        }

        /// <summary>
        /// true if SalesManager recieves a notice
        /// </summary>
        public bool SalesManager
        {
            get
            {
                return IsSalesManager;
            }
            set
            {
                IsSalesManager = value;
            }
        }

        /// <summary>
        /// true if Estimator recieves a notice
        /// </summary>
        public bool Estimator
        {
            get
            {
                return IsEstimator;
            }
            set
            {
                IsEstimator = value;
            }
        }

        /// <summary>
        /// true if VendorPurchasing recieves a notice
        /// </summary>
        public bool VendorPurchasing
        {
            get
            {
                return IsVendorPurchasing;
            }
            set
            {
                IsVendorPurchasing = value;
            }
        }

        /// <summary>
        /// true if VendorScheduling recieves a notice
        /// </summary>
        public bool VendorScheduling
        {
            get
            {
                return IsVendorScheduling;
            }
            set
            {
                IsVendorScheduling = value;
            }
        }

        /// <summary>
        /// true if VendorService recieves a notice
        /// </summary>
        public bool VendorService
        {
            get
            {
                return IsVendorService;
            }
            set
            {
                IsVendorService = value;
            }
        }

        /// <summary>
        /// true if Purchaser recieves a notice
        /// </summary>
        public bool Purchaser
        {
            get
            {
                return IsPurchaser;
            }
            set
            {
                IsPurchaser = value;
            }
        }

        /// <summary>
        /// true if DCSalesPerson recieves a notice
        /// </summary>
        public bool DCSalesPerson
        {
            get
            {
                return IsDCSalesPerson;
            }
            set
            {
                IsDCSalesPerson = value;
            }
        }

        /// <summary>
        /// true if AccountingAdmin recieves a notice
        /// </summary>
        public bool AccountingAdmin
        {
            get
            {
                return IsAccountingAdmin;
            }
            set
            {
                IsAccountingAdmin = value;
            }
        }

        /// <summary>
        /// true if SystemAdmin recieves a notice
        /// </summary>
        public bool SystemAdmin
        {
            get
            {
                return IsSystemAdmin;
            }
            set
            {
                IsSystemAdmin = value;
            }
        }

        /// <summary>
        /// true if Lender recieves a notice
        /// </summary>
        public bool Lender
        {
            get
            {
                return IsLender;
            }
            set
            {
                IsLender = value;
            }
        }

        /// <summary>
        /// true if TitleCompany recieves a notice
        /// </summary>
        public bool TitleCompany
        {
            get
            {
                return IsTitleCompany;
            }
            set
            {
                IsTitleCompany = value;
            }
        }

        /// <summary>
        /// true if Realtor recieves a notice
        /// </summary>
        public bool Realtor
        {
            get
            {
                return IsRealtor;
            }
            set
            {
                IsRealtor = value;
            }
        }

        /// <summary>
        /// true if Broker recieves a notice
        /// </summary>
        public bool Broker
        {
            get
            {
                return IsBroker;
            }
            set
            {
                IsBroker = value;
            }
        }

        /// <summary>
        /// true if SellingRealtor recieves a notice
        /// </summary>
        public bool SellingRealtor
        {
            get
            {
                return IsSellingRealtor;
            }
            set
            {
                IsSellingRealtor = value;
            }
        }

        /// <summary>
        /// Sets/Gets trigger name
        /// </summary>
        public string TriggerName
        {
            get
            {
                return strTriggerName;
            }
            set
            {
                strTriggerName = value;
            }
        }

        /// <summary>
        /// Sets/Gets recipients
        /// </summary>
        public string Recipient
        {
            get
            {
                return strRecipient;
            }
            set
            {
                strRecipient = value;
            }
        }
        public string CC
        {
            get
            {
                return strCC;
            }
            set
            {
                strCC = value;
            }
        }
        /// <summary>
        /// Sets/Gets FromAddress
        /// </summary>
        public string FromAddress
        {
            get
            {
                return strFromAddress;
            }
            set
            {
                strFromAddress = value;
            }
        }

        /// <summary>
        /// Sets/Gets subject
        /// </summary>
        public string Subject
        {
            get
            {
                return strSubject;
            }
            set
            {
                strSubject = value;
            }
        }

        /// <summary>
        /// Sets/Gets message
        /// </summary>
        public string Message
        {
            get
            {
                return strMessage;
            }
            set
            {
                strMessage = value;
                
                this.webBrowser.Navigate("about:blank");
                HtmlDocument doc = this.webBrowser.Document;
                doc.Write(string.Empty);
                this.webBrowser.DocumentText = strMessage; 
            }
        }

        /// <summary>
        /// Sets/Gets attachment
        /// </summary>
        public string Attachment
        {
            get
            {
                return strAttachment;
            }
            set
            {
                strAttachment = value;
            }
        }

        /// <summary>
        /// Sets/Gets report
        /// </summary>
        public string Report
        {
            get
            {
                return strReport;
            }
            set
            {
                strReport = value;
            }
        }

        private void AutoEventDetails_Load(object sender, EventArgs e)
        {
            flexParams.Rows.Count = 1;
            IsLoading = true;

            LoadData();
       
            IsLoading = false;
            IsModified = false;
        }

             
        private void LoadData()
        {
            txtTrigger.Text = strTriggerName;
            txtRecipient.Text = strRecipient;
            txtCC.Text = strCC;

            txtSender.Text = strFromAddress;
            txtSubject.Text = strSubject;
            txtMessage.Text = strMessage;
            txtAttachment.Text = strAttachment;
            txtReport.Text = strReport;

            chkDisabled.Checked = IsDisabled;
            chkBuyer.Checked = IsBuyer;
            chkCoBuyer.Checked = IsCoBuyer;
            chkSalesPerson.Checked = IsSalesPerson;
            chkPM.Checked = IsPM;
            chkSalesManager.Checked= IsSalesManager;
            chkEstimator.Checked = IsEstimator;

            chkVendorPurchasing.Checked = IsVendorPurchasing;
            chkVendorScheduling.Checked = IsVendorScheduling;
            chkVendorService.Checked = IsVendorService;

            chkPurchaser.Checked = IsPurchaser;
            chkDCSalesPerson.Checked = IsDCSalesPerson;

            chkAccountingAdmin.Checked = IsAccountingAdmin;
            chkSystemAdmin.Checked = IsSystemAdmin;

            chkLender.Checked = IsLender;
            chkTitleCompany.Checked = IsTitleCompany;

            chkRealtor.Checked = IsRealtor;
            chkBroker.Checked = IsBroker;
            chkSellingRealtor.Checked = IsSellingRealtor;
        }

        /// <summary>
        /// Canceled button clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmdCancel_Click(object sender, EventArgs e)
        {
            IsCanceled = true;
            this.Close();
        }

        private void flexParams_DoubleClick(object sender, EventArgs e)
        {
            //TriggerParamDetails param = new TriggerParamDetails();
            //param.ShowDialog(this);
        }

        private void txtReport_TextChanged(object sender, EventArgs e)
        {
            if (IsLoading)
                LoadReportParamsWhenLoading();
            else
            {
                LoadReportParamsWhenReportChanged();
                IsModified = true;
            }
        }

        /// <summary>
        /// Returns report parameters collection
        /// </summary>
        /// <param name="rptFile"></param>
        private void GetReportParams(string rptFile)
        {
            ReportDocument CRReport = new ReportDocument();
            CRReport.Load(rptFile);
            CRReport.Refresh();

            renamedParams = 0;
            ParameterFields parameters = CRReport.ParameterFields;
            for (int i = 0; i < parameters.Count; i++)
            {
                if (IsLoading)
                {
                    if (!parameters[i].Name.StartsWith("pm", true, null))
                    {
                        ReportParam param = (ReportParam)reportParams[i];
                        param.Name = parameters[i].Name;
                        reportParams[i] = param;
                        renamedParams++;
                    }
                }
                else
                {
                    if (!parameters[i].Name.StartsWith("pm", true, null))
                    {
                        ReportParam param = new ReportParam();
                        param.Name = parameters[i].Name;
                        param.Index = i;
                        reportParams.Add(param);
                    }
                }
            }
        }

        private void picAttachment_Click(object sender, EventArgs e)
        {
            openAttachment.ShowDialog(this);
            this.txtAttachment.Text = openAttachment.FileName;
        }

        private void picReport_Click(object sender, EventArgs e)
        {
            openReport.ShowDialog(this);
            this.txtReport.Text = openReport.FileName;
        }

        private void cmdSave_Click(object sender, EventArgs e)
        {
            strTriggerName = txtTrigger.Text;
            strRecipient = txtRecipient.Text;
            strCC= txtCC.Text;
            strFromAddress = txtSender.Text;
            strSubject = txtSubject.Text;
            strMessage = txtMessage.Text;
            strAttachment = txtAttachment.Text;
            strReport = txtReport.Text;

            IsDisabled = chkDisabled.Checked;

            IsBuyer = chkBuyer.Checked;
            IsCoBuyer = chkCoBuyer.Checked;
            IsSalesPerson = chkSalesPerson.Checked;
            IsPM = chkPM.Checked;
            IsSalesManager = chkSalesManager.Checked;
            IsEstimator = chkEstimator.Checked;

            IsVendorPurchasing = chkVendorPurchasing.Checked;
            IsVendorScheduling = chkVendorScheduling.Checked;
            IsVendorService = chkVendorService.Checked;

            IsPurchaser = chkPurchaser.Checked;
            IsDCSalesPerson = chkDCSalesPerson.Checked;

            IsAccountingAdmin = chkAccountingAdmin.Checked;
            IsSystemAdmin = chkSystemAdmin.Checked;

            IsLender = chkLender.Checked;
            IsTitleCompany = chkTitleCompany.Checked;

            IsRealtor = chkRealtor.Checked;
            IsBroker = chkBroker.Checked;
            IsSellingRealtor = chkSellingRealtor.Checked;

           
            // save report params
            reportParams.Clear();
            for (int i = 1; i < flexParams.Rows.Count; i++)
            {
                ReportParam rep_param = new ReportParam();
                rep_param.Name = "Param" + (i - 1);
                rep_param.Index = Convert.ToInt32(flexParams[i,"Index"]);

                if (flexParams[i,"Value"] != null)
                    rep_param.Value = flexParams[i,"Value"].ToString();
                else
                    rep_param.Value = "";
                reportParams.Add(rep_param);
            }
            this.Close();
        }

        private string GetTriggerTable(string triggerName)
        {
            SqlConnection cn = new SqlConnection(sqlConnStr);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT dbo.AutoNotice_GetTriggerTable('" 
                            + triggerName + "')";

            cn.Open();
            try
            {
                string tbl_name = cmd.ExecuteScalar().ToString();
                cn.Close();
                return tbl_name;
            }
            catch (Exception ex)
            {
               MessageBox.Show(ex.Message, "Getting trigg()er table");
               cn.Close();
               return "";
            }          
        }

        private void txtTrigger_TextChanged(object sender, EventArgs e)
        {
            if (IsLoading)
            {
                strTriggerTableName = GetTriggerTable(txtTrigger.Text);
                return;
            }

            IsModified = true;
        }

        private void txtRecipient_TextChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void txtSubject_TextChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void txtMessage_TextChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void txtAttachment_TextChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void flexParams_CellChanged(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            IsModified = true;
        }

        private void chkBuyer_CheckedChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void chkCoBuyer_CheckedChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void chkSalesPerson_CheckedChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void chkDCSalesPerson_CheckedChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void chkSalesManager_CheckedChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void chkPM_CheckedChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void chkEstimator_CheckedChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void chkPurchaser_CheckedChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void chkVendorPurchasing_CheckedChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void chkVendorScheduling_CheckedChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void chkVendorService_CheckedChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void chkAccountingAdmin_CheckedChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void chkSystemAdmin_CheckedChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void chkTitleCompany_CheckedChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void chkLender_CheckedChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void chkRealtor_CheckedChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void chkBroker_CheckedChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void chkSellingRealtor_CheckedChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        /// <summary>
        /// Load report parameters when report has been changed
        /// </summary>
        private void LoadReportParamsWhenReportChanged()
        {
            reportParams.Clear();
            flexParams.Rows.Count = 1;
            
            if (string.IsNullOrEmpty(txtReport.Text))
            {
                flexParams.AllowEditing = false;
                return;
            }

            GetReportParams(txtReport.Text);
            if (reportParams.Count == 0)
            {
                flexParams.AllowEditing = false;
                return;
            }

            for (int i = 0; i < reportParams.Count; i++)
            {
                flexParams.Rows.Add();
                ReportParam rep_param = (ReportParam)reportParams[i];
                flexParams[i + 1, "Parameter"] = rep_param.Name;
                flexParams[i + 1, "Index"] = rep_param.Index;
            }
        }

        /// <summary>
        /// Load report parameters when form is being loaded
        /// </summary>
        private void LoadReportParamsWhenLoading()
        {
            if (string.IsNullOrEmpty(txtReport.Text))
            {
                flexParams.Rows.Count = 1;
                flexParams.AllowEditing = false;
                return;
            }

            GetReportParams(txtReport.Text);

            flexParams.Rows.Add(renamedParams);
            for (int i = 0; i < renamedParams; i++)
            {
                ReportParam rep_param = (ReportParam)reportParams[i];
                flexParams[i+1,"Parameter"] = rep_param.Name;
                flexParams[i+1,"Index"] = rep_param.Index;
                flexParams[i+1,"Value"] = rep_param.Value;
            }
        }

        
        private void picSubject_Click(object sender, EventArgs e)
        {
            int pos = txtSubject.SelectionStart;

            TriggerParamDetails param_details = new TriggerParamDetails();
            param_details.ConnectionString = sqlConnStr;
            param_details.TriggerTable = strTriggerTableName;
            param_details.Trigger = this.txtTrigger.Text;
            param_details.ShowDialog(this);

            string textToInsert = param_details.SelectedValue;
            string newText = txtSubject.Text.Substring(0, pos) +
                             textToInsert +
                             txtSubject.Text.Substring(pos, txtSubject.Text.Length - pos);

            txtSubject.Text = newText;
            txtSubject.SelectionStart = pos;

        }

        private void picMessage_Click(object sender, EventArgs e)
        {
            MessageEditor msg_editor = new MessageEditor();
            msg_editor.ConnectionString = sqlConnStr;
            msg_editor.Message = this.txtMessage.Text;
            msg_editor.TriggerTable = strTriggerTableName;
            msg_editor.Trigger = this.txtTrigger.Text;
            msg_editor.ShowDialog(this);

            if (msg_editor.Canceled) return;

            this.txtMessage.Text = msg_editor.Message;

            this.webBrowser.Navigate("about:blank");
            HtmlDocument doc = this.webBrowser.Document;
            doc.Write(string.Empty);
            this.webBrowser.DocumentText = msg_editor.Message; 
        }

        private void flexParams_CellButtonClick(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            TriggerParamDetails param_details = new TriggerParamDetails();
            param_details.ConnectionString = sqlConnStr;
            param_details.TriggerTable = strTriggerTableName;
            param_details.Trigger = this.txtTrigger.Text;
            param_details.IsReportParameter = true;
            param_details.ShowDialog(this);

            if (param_details.Canceled) return;

            flexParams[e.Row, e.Col] = param_details.SelectedValue;
        }

        private void chkDisabled_CheckedChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void txtSender_TextChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void flexParams_AfterEdit(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            switch (flexParams.Cols[e.Col].Name)
            {
                case "Value":
                    {
                        for (int i = e.Row; i <= flexParams.RowSel; i++)
                            flexParams[i, e.Col] = flexParams[e.Row, e.Col];
                        break;
                    }
            }
        }

        private void txtCC_TextChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }
    }
}
