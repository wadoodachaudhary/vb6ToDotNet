using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace AutoNotice
{
    public partial class MessageEditor : Form
    {
        /// <summary>
        /// Connection string
        /// </summary>
        private string sqlConnStr;

        /// <summary>
        /// Trigger Name
        /// </summary>
        private string strMessageText;

        /// <summary>
        /// 
        /// </summary>
        private bool IsCanceled;

        /// <summary>
        /// Table Name which Trigger is based on
        /// </summary>
        private string strTableName;

        /// <summary>
        /// Trigger Name
        /// </summary>
        private string strTriggerName;

        /// <summary>
        /// Recipient query
        /// </summary>
        private string strRecipientQuery;

        /// <summary>
        /// Selected Table Name
        /// </summary>
        public string TriggerTable
        {
            get
            {
                return strTableName;
            }
            set
            {
                strTableName = value;
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
        /// Selected Trigger
        /// </summary>
        public string Trigger
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
        /// Sets/Gets message text
        /// </summary>
        public string Message
        {
            get
            {
                return strMessageText;
            }
            set
            {
                strMessageText  = value;
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

        public MessageEditor()
        {
            InitializeComponent();
        }

        private void MessageEditor_Load(object sender, EventArgs e)
        {
            htmlEditor.Message = strMessageText;
            if (string.IsNullOrEmpty(strTriggerName))
            {
                if (string.IsNullOrEmpty(strRecipientQuery))
                    cmdInsertParams.Visible = false;
            }
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            IsCanceled = true;
            this.Close();
        }

        private void cmdSave_Click(object sender, EventArgs e)
        {
            strMessageText = htmlEditor.Message;
            this.Close();
        }

        private void cmdInsertParams_Click(object sender, EventArgs e)
        {
            TriggerParamDetails param_details = new TriggerParamDetails();
            param_details.ConnectionString = sqlConnStr;
            param_details.TriggerTable = strTableName;
            param_details.Trigger = strTriggerName;
            param_details.IsReportParameter = false;
            param_details.RecipientQuery = strRecipientQuery;
            param_details.ShowDialog(this);

            htmlEditor.InsertParameter(param_details.SelectedValue);
            htmlEditor.Focus();

        }

        private void htmlEditor_Load(object sender, EventArgs e)
        {

        }
    }
}
