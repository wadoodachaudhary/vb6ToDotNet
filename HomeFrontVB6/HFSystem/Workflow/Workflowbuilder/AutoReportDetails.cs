using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

namespace AutoNotice
{
    struct ReportParam
    {
        public string Name;
        public int Index;
        public string Value;
    }

    public partial class AutoReportDetails : Form
    {
        public AutoReportDetails()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Connection string
        /// </summary>
        private string sqlConnStr;

        /// <summary>
        /// ID of report
        /// </summary>
        private int Seq;

        /// <summary>
        /// Description
        /// </summary>
        private string strDescription;

        /// <summary>
        /// Recipients
        /// </summary>
        private string strRecipient;

        /// <summary>
        /// Recipient query
        /// </summary>
        private string strRecipientQuery;

        /// <summary>
        /// Sender
        /// </summary>
        private string strSender;

        /// <summary>
        /// Subject
        /// </summary>
        private string strSubject;

        /// <summary>
        /// Message
        /// </summary>
        private string strMessage;

        /// <summary>
        /// Attachment
        /// </summary>
        private string strAttachment;

        /// <summary>
        /// Report
        /// </summary>
        private string strReport;

        /// <summary>
        /// Frequency Type
        /// </summary>
        private string strFrequencyType;

        /// <summary>
        /// Frequency Value
        /// </summary>
        private int intFrequencyValue;

        /// <summary>
        /// Run Time
        /// </summary>
        private DateTime dtRunTime;

        /// <summary>
        /// Next Run
        /// </summary>
        private DateTime dtNextRun;

        string[] DaysOfWeek = new string[] { "Every Week Day", "Sunday", "Monday", 
                    "Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"};

        string[] DaysOfMonth = new string[] { "Last Day", "1", "2","3","4","5","6",
                    "7","8","9","10","11","12","13","14","15","16","17","18","19","20",
                    "21","22","23","24","25","26","27","28","29","30","31"};
        
        string[] EveryNthWeek = new string[] {"1", "2","3","4","5"};
       
        string[] EveryNthMonth = new string[] {"1", "2","3","4","5","6","7","8","9","10","11","12"};

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
        /// Flag, if Cancel button has been clicked or not
        /// </summary>
        private bool IsCanceled;

        /// <summary>
        /// Flag, if data have been modified
        /// </summary>
        private bool IsModified;

        /// <summary>
        /// Flag, if description have been modified
        /// </summary>
        private bool IsDescriptionChanged;
              
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
        /// Report description
        /// </summary>
        public string Description
        {
            get
            {
                return strDescription;
            }
            set
            {
                strDescription = value;
            }
        }

        /// <summary>
        /// Sets/Gets Recipient Query
        /// </summary>
        public string RecipientQuery
        {
            get
            {
                return strRecipientQuery;
            }
            set
            {
                strRecipientQuery = value;
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
        /// Sets/Gets connection string
        /// </summary>
        public int ReportID
        {
            get
            {
                return Seq;
            }
            set
            {
                Seq = value;
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

        /// <summary>
        /// Sets/Gets sender
        /// </summary>
        public string Sender
        {
            get
            {
                return strSender;
            }
            set
            {
                strSender = value;
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

                this.webSubject.Navigate("about:blank");
                HtmlDocument doc = this.webSubject.Document;
                doc.Write(string.Empty);
                this.webSubject.DocumentText = strSubject;
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

        /// <summary>
        /// Sets/Gets Frequency Type
        /// </summary>
        public string FrequencyType
        {
            get
            {
                return strFrequencyType;
            }
            set
            {
                strFrequencyType = value;
            }
        }

        /// <summary>
        /// Sets/Gets Frequency Value
        /// </summary>
        public int FrequencyValue
        {
            get
            {
                return intFrequencyValue;
            }
            set
            {
                intFrequencyValue = value;
            }
        }

        /// <summary>
        /// Sets/Gets Run Time
        /// </summary>
        public DateTime RunTime
        {
            get
            {
                return dtRunTime;
            }
            set
            {
                dtRunTime = value;
            }
        }

        /// <summary>
        /// Sets/Gets Next Run
        /// </summary>
        public DateTime NextRun
        {
            get
            {
                return dtNextRun;
            }
            set
            {
                dtNextRun = value;
            }
        }

        private void SetupGrid()
        {
            for (int i = 0; i < 10; i++)
            {
                
                flexParams[i + 1, "Parameter"] = "Param" + (i + 1);
                flexParams[i + 1, "Index"] = i;
                flexParams[i + 1, "Value"] = "";
            }
        }

        private void AutoReportDetails_Load(object sender, EventArgs e)
        {
            SetupGrid();
            IsLoading = true;
           
            LoadData();

            IsLoading = false;
            IsModified = false;
            IsDescriptionChanged = false;
        }

        private void LoadData()
        {
            if (Seq == -1)
            {
                cmbFrequencyType.SelectedItem = "Day of week";
                cmbFrequencyValue.SelectedItem = cmbFrequencyValue.Items[0];
                return;
            }

            txtDescription.Text = strDescription;
            txtRecipient.Text = strRecipient;
            txtSender.Text = strSender;
            txtSubject.Text = strSubject;
            txtMessage.Text = strMessage;
            txtAttachment.Text = strAttachment;
            txtReport.Text = strReport;
            txtRecQuery.Text = strRecipientQuery;
            IsDescriptionChanged = false;


            if (string.IsNullOrEmpty(txtReport.Text))
                LoadReportParamsWhenLoading();

            if (!string.IsNullOrEmpty(strRecipientQuery))
                flexParams.Cols["Value"].ComboList = "|...";

            runTime.Value = dtRunTime;
            nextRun.Value = dtNextRun;

            if (!string.IsNullOrEmpty(strFrequencyType))
                cmbFrequencyType.SelectedItem = strFrequencyType;
            else
                cmbFrequencyType.SelectedItem = "Day of week";

            if ((strFrequencyType == "Day of week")
                || (strFrequencyType == "Day of month"))
                cmbFrequencyValue.SelectedIndex = intFrequencyValue;
            else
                cmbFrequencyValue.SelectedItem = intFrequencyValue.ToString();
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            IsCanceled = true;
            this.Close();
        }

        private void cmbFrequencyType_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (cmbFrequencyType.Text)
            {
                case "Day of week":
                    {
                        cmbFrequencyValue.Items.Clear();
                        cmbFrequencyValue.Items.AddRange(DaysOfWeek);
                        cmbFrequencyValue.Text = cmbFrequencyValue.Items[0].ToString();
                        break;
                    }
                case "Day of month":
                    {
                        cmbFrequencyValue.Items.Clear();
                        cmbFrequencyValue.Items.AddRange(DaysOfMonth);
                        cmbFrequencyValue.Text = cmbFrequencyValue.Items[0].ToString();
                        break;
                    }
                case "Every nth month":
                    {
                        cmbFrequencyValue.Items.Clear();
                        cmbFrequencyValue.Items.AddRange(EveryNthMonth);
                        cmbFrequencyValue.Text = cmbFrequencyValue.Items[0].ToString();
                        break;
                    }
                case "Every nth week":
                    {
                        cmbFrequencyValue.Items.Clear();
                        cmbFrequencyValue.Items.AddRange(EveryNthWeek);
                        cmbFrequencyValue.Text = cmbFrequencyValue.Items[0].ToString();
                        break;
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

        /// <summary>
        /// Save form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmdSave_Click(object sender, EventArgs e)
        {
            if (!IsReportValid()) return;

            strDescription = txtDescription.Text;
            strRecipient = txtRecipient.Text;
            strSender = txtSender.Text;
            strSubject = txtSubject.Text;
            strMessage = txtMessage.Text;
            strAttachment = txtAttachment.Text;
            strReport = txtReport.Text;

            strRecipientQuery = txtRecQuery.Text;

            dtRunTime = runTime.Value;
            dtNextRun = nextRun.Value;

            strFrequencyType = cmbFrequencyType.Text;
            
            if ((strFrequencyType == "Day of week")
                ||(strFrequencyType == "Day of month"))
                intFrequencyValue = cmbFrequencyValue.SelectedIndex;
            else
                intFrequencyValue = Convert.ToInt32(cmbFrequencyValue.Text);

            try
            {
                // save report params
                reportParams.Clear();
                for (int i = 1; i < flexParams.Rows.Count; i++)
                {
                    ReportParam rep_param = new ReportParam();
                    rep_param.Name = "Param" + i;
                    rep_param.Index = Convert.ToInt32(flexParams[i, "Index"]);

                    if (flexParams[i, "Value"] != null)
                        rep_param.Value = flexParams[i, "Value"].ToString();
                    else
                        rep_param.Value = "";
                    reportParams.Add(rep_param);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Report Save Error");
            }
            this.Close();
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

        /// <summary>
        /// Load report parameters when form is being loaded
        /// </summary>
        private void LoadReportParamsWhenLoading()
        {
            if (!string.IsNullOrEmpty(txtReport.Text))
                GetReportParams(txtReport.Text);

            for (int i = 0; i < reportParams.Count; i++)
            {
                ReportParam rep_param = (ReportParam)reportParams[i];
                flexParams[i + 1, "Value"] = rep_param.Value;
            }

            for (int i = 0; i < renamedParams; i++)
            {
                ReportParam rep_param = (ReportParam)reportParams[i];
                flexParams[i + 1, "Parameter"] = flexParams[i + 1, "Parameter"].ToString() +
                                                "(" + rep_param.Name + ")";
            }
        }

        /// <summary>
        /// Load report parameters when report has been changed
        /// </summary>
        private void LoadReportParamsWhenReportChanged()
        {
            reportParams.Clear();
            for (int i = 0; i < flexParams.Rows.Count - 1; i++)
            {
                flexParams[i + 1, "Parameter"] = "Param" + (i + 1);
                flexParams[i + 1, "Value"] = "";
            }

            if (string.IsNullOrEmpty(txtReport.Text)) return;

            GetReportParams(txtReport.Text);
            if (reportParams.Count == 0) return;

            for (int i = 0; i < reportParams.Count; i++)
            {
                ReportParam rep_param = (ReportParam)reportParams[i];
                flexParams[i + 1, "Parameter"] = flexParams[i + 1, "Parameter"].ToString() +
                                                "(" + rep_param.Name + ")";
                flexParams[i + 1, "Index"] = rep_param.Index;
                flexParams[i + 1, "Value"] = rep_param.Value; 
            }        
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

        private void txtRecipient_TextChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void txtSender_TextChanged(object sender, EventArgs e)
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

        private void grvParams_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            IsModified = true;
        }

        private void runTime_ValueChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void nextRun_ValueChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void cmbFrequencyType_SelectedValueChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void cmbFrequencyValue_TextChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void cmbFrequencyValue_SelectedValueChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void picMessage_Click(object sender, EventArgs e)
        {
            MessageEditor msg_editor = new MessageEditor();
            msg_editor.ConnectionString = sqlConnStr;
            msg_editor.Message = this.txtMessage.Text;
            msg_editor.RecipientQuery = txtRecQuery.Text;
            msg_editor.ShowDialog(this);

            if (msg_editor.Canceled) return;

            this.txtMessage.Text = msg_editor.Message;

            this.webBrowser.Navigate("about:blank");
            HtmlDocument doc = this.webBrowser.Document;
            doc.Write(string.Empty);
            this.webBrowser.DocumentText = msg_editor.Message; 

        }

        private void cmdQuery_Click(object sender, EventArgs e)
        {
            Query query = new Query();
            query.ConnectionString = sqlConnStr;
            query.RecipientQuery = txtRecQuery.Text;
            query.ShowDialog(this);

            if (query.Saved)
                txtRecQuery.Text = query.RecipientQuery;

            if (!string.IsNullOrEmpty(txtRecQuery.Text))
                flexParams.Cols["Value"].ComboList = "|...";
            else
                flexParams.Cols["Value"].ComboList = "";
        }

        private void txtRecQuery_TextChanged(object sender, EventArgs e)
        {
            IsModified = true;
        }

        private void picSubject_Click(object sender, EventArgs e)
        {
            MessageEditor msg_editor = new MessageEditor();
            msg_editor.ConnectionString = sqlConnStr;
            msg_editor.Message = this.txtSubject.Text;
            msg_editor.RecipientQuery = txtRecQuery.Text;
            msg_editor.ShowDialog(this);

            if (msg_editor.Canceled) return;

            this.txtSubject.Text = msg_editor.Message;

            this.webSubject.Navigate("about:blank");
            HtmlDocument doc = this.webSubject.Document;
            doc.Write(string.Empty);
            this.webSubject.DocumentText = msg_editor.Message; 
        }

        private void txtDescription_TextChanged(object sender, EventArgs e)
        {
            IsModified = true;
            IsDescriptionChanged = true;
        }

        /// <summary>
        /// Checks if report description is already exists
        /// </summary>
        /// <param name="description"></param>
        /// <returns></returns>
        private bool IsDescriptionExists(string description)
        {
            SqlConnection cn = new SqlConnection(sqlConnStr);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT dbo.AutoNotice_IsReportDescriptionExists('" +
                                description + "')";
            cmd.Connection = cn;

            cn.Open();
            try
            {
                bool result = Convert.ToBoolean(cmd.ExecuteScalar());
                cn.Close();
                return result;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message,"Validation",MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                cn.Close();
                return false;
            }
        }

        /// <summary>
        /// Validates report
        /// </summary>
        /// <returns></returns>
        private bool IsReportValid()
        {
            if (string.IsNullOrEmpty(txtDescription.Text))
            {
                MessageBox.Show("Report description cannot be blank." +
                                "\r\n" + "Please, enter the description.",
                                "Validation", MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                return false;
            }

            if (IsDescriptionChanged)
            {
                if (IsDescriptionExists(txtDescription.Text))
                {
                    MessageBox.Show("Report with this description is already exists." +
                                    "\r\n" + "Please, select another description.",
                                    "Validation", MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                    return false;
                }
            }
            return true;
        }

        private void flexParams_CellChanged(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            IsModified = true;
        }

        private void flexParams_CellButtonClick(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            TriggerParamDetails param_details = new TriggerParamDetails();
            param_details.ConnectionString = sqlConnStr;
            param_details.TriggerTable = "";
            param_details.Trigger = "";
            param_details.IsReportParameter = true;
            param_details.RecipientQuery = txtRecQuery.Text; 
            param_details.ShowDialog(this);

            if (param_details.Canceled) return;
            flexParams[e.Row, e.Col] = param_details.SelectedValue;
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

        private void cmdTestMe_Click(object sender, EventArgs e)
        {
            TestMe testMe = new TestMe();
            testMe.ShowDialog(this);

            if (string.IsNullOrEmpty(testMe.Recipient)) return;

            TestMeNow(testMe.Recipient);
        }

        private void TestMeNow(string recipient)
        {
            SqlConnection cn = new SqlConnection(sqlConnStr);
            SqlCommand cmd = new SqlCommand();

            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "AutoNotice_TestMe";
            cmd.Connection = cn;

            cmd.Parameters.AddWithValue("@Description",txtDescription.Text);
            cmd.Parameters.AddWithValue("@RecipientAddress",recipient);
            cmd.Parameters.AddWithValue("@FromAddress",txtSender.Text);
            cmd.Parameters.AddWithValue("@Subject",txtSubject.Text);
            cmd.Parameters.AddWithValue("@Message",txtMessage.Text);

            cmd.Parameters.AddWithValue("@Attachment",txtAttachment.Text);
            cmd.Parameters.AddWithValue("@Report",txtReport.Text);
           
            cmd.Parameters.AddWithValue("@Param1",flexParams[1,"Value"].ToString());
            cmd.Parameters.AddWithValue("@Param2",flexParams[2,"Value"].ToString());
            cmd.Parameters.AddWithValue("@Param3",flexParams[3,"Value"].ToString());

            cmd.Parameters.AddWithValue("@Param4",flexParams[4,"Value"].ToString());
            cmd.Parameters.AddWithValue("@Param5",flexParams[5,"Value"].ToString());
            cmd.Parameters.AddWithValue("@Param6",flexParams[6,"Value"].ToString());
            cmd.Parameters.AddWithValue("@Param7",flexParams[7,"Value"].ToString());
            cmd.Parameters.AddWithValue("@Param8",flexParams[8,"Value"].ToString());

            cmd.Parameters.AddWithValue("@Param9",flexParams[9,"Value"].ToString());
            cmd.Parameters.AddWithValue("@Param10", flexParams[10,"Value"].ToString());
           
            cn.Open();
            try
            {
                cmd.ExecuteNonQuery();
                cn.Close();

                MessageBox.Show("Test was successfull.", "Test Me", 
                    MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Test Me", MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                cn.Close();
            }
        }
                        
    }
}
