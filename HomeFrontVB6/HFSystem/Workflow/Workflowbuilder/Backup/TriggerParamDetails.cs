using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Data.SqlClient;

namespace AutoNotice
{
    public partial class TriggerParamDetails : Form
    {
        public TriggerParamDetails()
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
        private string strSelectedValue;

        /// <summary>
        /// Table Name
        /// </summary>
        private string strTableName;

        /// <summary>
        /// Trigger Name
        /// </summary>
        private string strTriggerName;

        /// <summary>
        /// true,if setting report parameter value
        /// </summary>
        private bool IsReportParam;

        /// <summary>
        /// true if Cancel button has been clicked
        /// </summary>
        private bool IsCanceled;

        /// <summary>
        /// Recipient query
        /// </summary>
        private string strRecipientQuery;

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
        /// Flag,if set a value for report parameter or not
        /// </summary>
        public bool IsReportParameter
        {
            get
            {
                return IsReportParam;
            }
            set
            {
                IsReportParam = value;
            }
        }

        /// <summary>
        /// Selected param value
        /// </summary>
        public string SelectedValue
        {
            get
            {
                return strSelectedValue;
            }
        }


        private void cmdCancel_Click(object sender, EventArgs e)
        {
            IsCanceled = true;
            this.Close();
        }

        private void OnSelect()
        {
            if (IsReportParam)
            {
                if (txtValue.Text != "")
                    strSelectedValue = txtValue.Text;
                else
                    strSelectedValue = grvTable.CurrentRow.Cells["FieldName"].Value.ToString();
            }
            else
            {
                if ((grvTable.CurrentRow.Cells["FieldType"].Value.ToString() == "datetime")
                    || (grvTable.CurrentRow.Cells["FieldType"].Value.ToString() == "smalldatetime"))
                {
                    strSelectedValue = "{{" + "D{" + grvTable.CurrentRow.Cells["FieldName"].Value.ToString() + "}D" + "}}";
                }
                else
                    if ((grvTable.CurrentRow.Cells["FieldType"].Value.ToString() == "int")
                        || (grvTable.CurrentRow.Cells["FieldType"].Value.ToString() == "bigint")
                        || (grvTable.CurrentRow.Cells["FieldType"].Value.ToString() == "tinyint")
                        || (grvTable.CurrentRow.Cells["FieldType"].Value.ToString() == "smallint")
                        || (grvTable.CurrentRow.Cells["FieldType"].Value.ToString() == "decimal")
                        || (grvTable.CurrentRow.Cells["FieldType"].Value.ToString() == "numeric")
                        || (grvTable.CurrentRow.Cells["FieldType"].Value.ToString() == "money")
                        || (grvTable.CurrentRow.Cells["FieldType"].Value.ToString() == "smallmoney")
                        || (grvTable.CurrentRow.Cells["FieldType"].Value.ToString() == "float")
                        || (grvTable.CurrentRow.Cells["FieldType"].Value.ToString() == "bit")
                        || (grvTable.CurrentRow.Cells["FieldType"].Value.ToString() == "int32")
                        || (grvTable.CurrentRow.Cells["FieldType"].Value.ToString() == "int16")
                        || (grvTable.CurrentRow.Cells["FieldType"].Value.ToString() == "double")
                        || (grvTable.CurrentRow.Cells["FieldType"].Value.ToString() == "bool"))
                    {
                        strSelectedValue = "{{" + "N{" + grvTable.CurrentRow.Cells["FieldName"].Value.ToString() + "}N" + "}}";
                    }
                    else
                        strSelectedValue = "{{" + grvTable.CurrentRow.Cells["FieldName"].Value.ToString() + "}}";
            }
            this.Close();
        }

        private void cmdOK_Click(object sender, EventArgs e)
        {
            OnSelect();
        }

        private void TriggerParamDetails_Load(object sender, EventArgs e)
        {
            this.sqlConn.ConnectionString = sqlConnStr;

            if (string.IsNullOrEmpty(strRecipientQuery))
            {
                this.sqlDATrigParams.SelectCommand.Parameters["@TableName"].Value = strTableName;
                this.sqlDATrigParams.SelectCommand.Parameters["@TriggerName"].Value = strTriggerName;
                this.sqlDATrigParams.Fill(dsTriggerParams);
            }
            else
            {
                this.sqlDATrigParams.SelectCommand.CommandType = CommandType.Text;
                this.sqlDATrigParams.SelectCommand.CommandText = strRecipientQuery;
                this.sqlDATrigParams.SelectCommand.Parameters.Clear();
                dsTriggerParams.Tables[0].Columns.Clear();

                grvTable.DataSource = null;

                this.sqlDATrigParams.Fill(dsTriggerParams);

                grvTable.Rows.Add(dsTriggerParams.Tables[0].Columns.Count);
                for (int i = 0; i < dsTriggerParams.Tables[0].Columns.Count; i++)
                {
                    grvTable.Rows[i].Cells["FieldName"].Value = dsTriggerParams.Tables[0].Columns[i].ColumnName;
                    grvTable.Rows[i].Cells["FieldType"].Value = dsTriggerParams.Tables[0].Columns[i].DataType.Name.ToLower();
                }
            }
                     

            //if (IsReportParam)
            //{
            //    lblTable.Text = "Select value from the list or enter in the text box.";
            //    lblParamValue.Visible = true;
            //    txtValue.Visible = true;
            //}
        }

        private void grvTable_DoubleClick(object sender, EventArgs e)
        {
            OnSelect();
        }

        private SqlCommand Search_Param(string paramName)
        {         
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = "AutoNotice_SearchTriggerParameter";
            cmd.CommandTimeout = 0;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Connection = this.sqlConn;
            
            cmd.Parameters.Add(new SqlParameter("@TableName",
                      SqlDbType.VarChar, 100));
            cmd.Parameters.Add(new SqlParameter("@TriggerName",
                      SqlDbType.VarChar, 100));
            cmd.Parameters.Add(new SqlParameter("@ParamName",
                      SqlDbType.VarChar, 100));
            cmd.Parameters["@TableName"].Value = strTableName;
            cmd.Parameters["@TriggerName"].Value = strTriggerName;
            cmd.Parameters["@ParamName"].Value = paramName; 

            return cmd;
        }

        private void cmdSearch_Click(object sender, EventArgs e)
        {
            dsTriggerParams.Clear();
            this.sqlDATrigParams.SelectCommand = Search_Param(this.txtParamName.Text);
            this.sqlDATrigParams.Fill(dsTriggerParams);            
        }

        private void txtParamName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                cmdSearch_Click(sender, e);
            }
        }

        private void grvTable_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                OnSelect();
        }

    }
}
