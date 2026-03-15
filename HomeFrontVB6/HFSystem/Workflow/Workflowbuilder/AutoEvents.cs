using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using C1.Win.C1FlexGrid;

namespace AutoNotice
{
    public partial class AutoEvents : Form
    {
        public AutoEvents()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 
        /// </summary>
        private bool IsModified;

        /// <summary>
        /// Connection String
        /// </summary>
        private string sqlConnStr;

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

        // Connection string
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


        private void AutoEvents_Load(object sender, EventArgs e)
        {
            this.sqlConn.ConnectionString = sqlConnStr;
            this.sqlDAEventsList.Fill(dsAutoEvents);

            SetupAutoEventsList();
            PopulateAutoEventsList();

            IsModified = false;
            flex.Focus();
        }

        /// <summary>
        /// Setup AutoReports list
        /// </summary>
        private void SetupAutoEventsList()
        {
            flex.Rows.Count = 1;
            flex.Cols.Count = dsAutoEvents.Tables["AutoNotice_GetAutoEventsList"].Columns.Count + 1;
            flex.ExtendLastCol = true;
        }

        /// <summary>
        /// Populates reports list
        /// </summary>
        private void PopulateAutoEventsList()
        {
            string name;
            flex.Rows.Count = 1;
            int row = 1;

            while (row <= dsAutoEvents.Tables["AutoNotice_GetAutoEventsList"].Rows.Count)
            {
                flex.Rows.Add();
                for (int i = 1; i < flex.Cols.Count; i++)
                {
                    name = flex.Cols[i].Name;
                    flex[row, flex.Cols[name].Index] = dsAutoEvents.Tables["AutoNotice_GetAutoEventsList"].Rows[row - 1][name];
                }
                row = flex.Rows.Count;
            }
        }

        
        private void flex_DoubleClick(object sender, EventArgs e)
        {
            OnEdit();
        }
        
        private void OnEdit()
        {
            HitTestInfo info = flex.HitTest();
            if (info.Type == HitTestTypeEnum.ColumnHeader) return;

            int row = flex.Row;

            AutoEventDetails trigger_details = new AutoEventDetails();
            trigger_details.DSN_NAME = DataSourceName;
            trigger_details.DB_USER = UID;
            trigger_details.DB_PASSWORD = PWD;
            trigger_details.ConnectionString = sqlConnStr;

            trigger_details.TriggerName = flex[row, "TriggerName"].ToString();
            trigger_details.Disabled = Convert.ToBoolean(flex[row, "IsDisabled"]);

            trigger_details.FromAddress = flex[row, "FromAddress"].ToString();
            trigger_details.Recipient = flex[row, "StaticRecipients"].ToString();
            trigger_details.CC = flex[row, "StaticCCs"].ToString();

            trigger_details.Subject = flex[row, "Subject"].ToString();
            trigger_details.Message = flex[row, "Message"].ToString();
            trigger_details.Attachment = flex[row, "Attachment"].ToString();
            trigger_details.Report = flex[row, "Report"].ToString();

            trigger_details.Buyer = Convert.ToBoolean(flex[row, "Buyer"]);
            trigger_details.CoBuyer = Convert.ToBoolean(flex[row, "CoBuyer"]);
            trigger_details.SalesPerson = Convert.ToBoolean(flex[row, "SalesPerson"]);
            trigger_details.PM = Convert.ToBoolean(flex[row, "PM"]);

            trigger_details.SalesManager = Convert.ToBoolean(flex[row, "SalesManager"]);
            trigger_details.Estimator = Convert.ToBoolean(flex[row, "Estimator"]);
            trigger_details.VendorPurchasing = Convert.ToBoolean(flex[row, "VendorPurchasing"]);
            trigger_details.VendorScheduling = Convert.ToBoolean(flex[row, "VendorScheduling"]);
            trigger_details.VendorService = Convert.ToBoolean(flex[row, "VendorService"]);

            trigger_details.Purchaser = Convert.ToBoolean(flex[row, "Purchaser"]);
            trigger_details.DCSalesPerson = Convert.ToBoolean(flex[row, "DCSalesPerson"]);
            trigger_details.AccountingAdmin = Convert.ToBoolean(flex[row, "AccountingAdministrator"]);
            trigger_details.SystemAdmin = Convert.ToBoolean(flex[row, "SystemAdministrator"]);

            trigger_details.Lender = Convert.ToBoolean(flex[row, "Lender"]);
            trigger_details.TitleCompany = Convert.ToBoolean(flex[row, "TitleCompany"]);
            trigger_details.Realtor = Convert.ToBoolean(flex[row, "Realtor"]);
            trigger_details.Broker = Convert.ToBoolean(flex[row, "Broker"]);
            trigger_details.SellingRealtor = Convert.ToBoolean(flex[row, "SellingRealtor"]);

            if (!string.IsNullOrEmpty(trigger_details.Report))
            {
                // set report parameters
                ArrayList rep_params = new ArrayList();
                for (int i = 1; i < 10; i++)
                {
                    ReportParam param = new ReportParam();
                    param.Name = "Param" + i;
                    param.Index = i - 1;
                    param.Value = (flex[row, "Param" + i] != null)
                                ? flex[row, "Param" + i].ToString()
                                : "";
                    rep_params.Add(param);
                }
                trigger_details.ReportParameters = rep_params;
            }

            trigger_details.ShowDialog(this);

            if ((trigger_details.Canceled) || (!trigger_details.Modified))
            {
                flex.AllowEditing = true;
                return;
            }

            flex[row, "TriggerName"] = trigger_details.TriggerName;
            flex[row, "IsDisabled"] = trigger_details.Disabled;

            flex[row, "FromAddress"] = trigger_details.FromAddress;
            flex[row, "StaticRecipients"] = trigger_details.Recipient;
            flex[row, "StaticCCs"] = trigger_details.CC;

            flex[row, "Subject"] = trigger_details.Subject;
            flex[row, "Message"] = trigger_details.Message;
            flex[row, "Attachment"] = trigger_details.Attachment;
            flex[row, "Report"] = trigger_details.Report;

            flex[row, "Buyer"] = trigger_details.Buyer;
            flex[row, "CoBuyer"] = trigger_details.CoBuyer;
            flex[row, "SalesPerson"] = trigger_details.SalesPerson;
            flex[row, "PM"] = trigger_details.PM;

            flex[row, "SalesManager"] = trigger_details.SalesManager;
            flex[row, "Estimator"] = trigger_details.Estimator;
            flex[row, "VendorPurchasing"] = trigger_details.VendorPurchasing;
            flex[row, "VendorScheduling"] = trigger_details.VendorScheduling;
            flex[row, "VendorService"] = trigger_details.VendorService;

            flex[row, "Purchaser"] = trigger_details.Purchaser;
            flex[row, "DCSalesPerson"] = trigger_details.DCSalesPerson;
            flex[row, "AccountingAdministrator"] = trigger_details.AccountingAdmin;
            flex[row, "SystemAdministrator"] = trigger_details.SystemAdmin;

            flex[row, "Lender"] = trigger_details.Lender;
            flex[row, "TitleCompany"] = trigger_details.TitleCompany;
            flex[row, "Realtor"] = trigger_details.Realtor;
            flex[row, "Broker"] = trigger_details.Broker;
            flex[row, "SellingRealtor"] = trigger_details.SellingRealtor;

            if (!string.IsNullOrEmpty(trigger_details.Report))
            {
                // get report parameters
                ArrayList rep_params = trigger_details.ReportParameters;
                for (int i = 0; i < rep_params.Count; i++)
                {
                    ReportParam param = (ReportParam)rep_params[i];
                    flex[row, "Param" + (param.Index + 1)] = param.Value;
                }
            }

            // format subject and message
            if (!string.IsNullOrEmpty(trigger_details.Subject))
            {
                string formattedSubject = trigger_details.Subject.Replace("'", "''''");
                formattedSubject = formattedSubject.Replace("D{", "ISNULL(CONVERT(VARCHAR(12),");
                formattedSubject = formattedSubject.Replace("}D", ",109),'')");

                formattedSubject = formattedSubject.Replace("N{", "ISNULL(CAST(CAST(round(");
                formattedSubject = formattedSubject.Replace("}N", ",2) AS money) AS varchar),'')");

                formattedSubject = formattedSubject.Replace("{{", "'+");
                formattedSubject = formattedSubject.Replace("}}", "+'");
                formattedSubject = "'" + formattedSubject + "'";
                flex[row, "FormattedSubject"] = formattedSubject;
            }

            if (!string.IsNullOrEmpty(trigger_details.Message))
            {
                string formattedMessage = trigger_details.Message.Replace("'", "''''");
                formattedMessage = formattedMessage.Replace("D{", "ISNULL(CONVERT(VARCHAR(12),");
                formattedMessage = formattedMessage.Replace("}D", ",109),'')");

                formattedMessage = formattedMessage.Replace("N{", "ISNULL(CAST(CAST(round(");
                formattedMessage = formattedMessage.Replace("}N", ",2) AS money) AS varchar),'')");

                formattedMessage = formattedMessage.Replace("{{", "'+");
                formattedMessage = formattedMessage.Replace("}}", "+'");
                formattedMessage = "'" + formattedMessage + "'";
                flex[row, "FormattedMessage"] = formattedMessage;
            }

            flex.AllowEditing = true;
            IsModified = true;
        }

        private void flex_BeforeDoubleClick(object sender, C1.Win.C1FlexGrid.BeforeMouseDownEventArgs e)
        {
            flex.AllowEditing = false;
        }

        private void cmdSave_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void Save()
        {
            //if (!IsValidated()) return;
            SaveToDataset();
            SaveToDatabase();
            IsModified = false;
        }

        
        /// <summary>
        /// Saves data from the grid to the dataset
        /// </summary>
        private void SaveToDataset()
        {
            string name;
            for (int i = 1; i < flex.Rows.Count; i++)
            {
                DataRow row = dsAutoEvents.Tables["AutoNotice_GetAutoEventsList"].Rows.Find(flex[i, flex.Cols["TriggerName"].Index]);

                for (int j = 2; j < flex.Cols.Count; j++)
                {
                    name = flex.Cols[j].Name;
                    if ((flex[i, flex.Cols[name].Index] != null)
                        && ((flex[i, flex.Cols[name].Index].ToString() != "")))
                        row[name] = flex[i, flex.Cols[name].Index];
                    else if (dsAutoEvents.Tables["AutoNotice_GetAutoEventsList"].Columns[name].AllowDBNull)
                        row[name] = Convert.DBNull;
                    else
                        row[name] = false;
                }
            }
        }

        private void SaveToDatabase()
        {
            try
            {
                this.sqlDAEventsList.Update(dsAutoEvents);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Saving events", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Deletes record from dataset
        /// </summary>
        /// <param name="UserID"></param>
        private void DeleteFromDataset(string triggerName)
        {
            for (int i = 0; i < dsAutoEvents.Tables["AutoNotice_GetAutoEventsList"].Rows.Count; i++)
            {
                DataRow row = dsAutoEvents.Tables["AutoNotice_GetAutoEventsList"].Rows[i];
                if (row.RowState != DataRowState.Deleted)
                {
                    if ((row["TriggerName"].ToString() == triggerName)
                        && (row.RowState != DataRowState.Deleted))
                    {
                        row.Delete();
                        return;
                    }
                }
            }

        }

        
        private void cmdClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void AutoEvents_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (IsModified)
            {
                DialogResult res = MessageBox.Show("Data have been modified. Would you like to save ?",
                                 "Saving", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    //if (!IsValidated())
                    //{
                    //    e.Cancel = true;
                    //    return;
                    //}
                    Save();
                }
            }
        }

        private void flex_AfterEdit(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            IsModified = true;
        }

        private void flex_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
                OnEdit();
        }

        private void flex_Click(object sender, EventArgs e)
        {

        }
    }
}
