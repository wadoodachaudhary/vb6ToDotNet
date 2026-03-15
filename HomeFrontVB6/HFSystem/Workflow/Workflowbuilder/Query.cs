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
    public partial class Query : Form
    {
        public Query()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Connection string
        /// </summary>
        private string sqlConnStr;

        /// <summary>
        /// Connection string
        /// </summary>
        private string strQueryText;

        /// <summary>
        /// Flag, if data have been modified
        /// </summary>
        private bool IsModified;

        /// <summary>
        /// Flag, if query has been saved
        /// </summary>
        private bool IsSaved;

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
        /// Sets/Gets connection string
        /// </summary>
        public string RecipientQuery
        {
            get
            {
                return strQueryText;
            }
            set
            {
                strQueryText = value;
            }
        }

        /// <summary>
        /// true if query has been saved
        /// </summary>
        public bool Saved
        {
            get
            {
                return IsSaved;
            }
        }

        private void txtQuery_TextChanged(object sender, EventArgs e)
        {
            IsModified = true;
            IsSaved = false; ;
        }

        private bool IsQueryValid()
        {
            string queryText = txtQuery.Text.Trim();
            if (queryText.Contains("RecipientAddress")
                || (string.IsNullOrEmpty(queryText)))
                return true;
            else
                return false;
        }

        private void RunQuery()
        {
            if (txtQuery.Text.Trim() == "") return;

            if (!IsQueryValid())
            {
                MessageBox.Show("Query must have RecipientAddress field.", "Validation !", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            SqlConnection cn = new SqlConnection(sqlConnStr);
            this.sqlQuery.SelectCommand.Connection = cn;
            this.sqlQuery.SelectCommand.CommandType = CommandType.Text;
            this.sqlQuery.SelectCommand.CommandText = txtQuery.Text.Trim();

            try
            {
                DataSet dsQuery = new DataSet();
                this.sqlQuery.Fill(dsQuery);

                grResult.DataSource = dsQuery.Tables[0]; ;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error !", MessageBoxButtons.OK, 
                    MessageBoxIcon.Error);
            }
        }

        private void mnuRunQuery_Click(object sender, EventArgs e)
        {
            RunQuery();
        }

        private void mnuExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Saves query
        /// </summary>
        private void Save()
        {
            if (!IsQueryValid())
            {
                MessageBox.Show("Query must have RecipientAddress field.", "Validation !",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            strQueryText = txtQuery.Text;
            IsSaved = true;
            IsModified = false;
        }

        private void mnuSave_Click(object sender, EventArgs e)
        {       
            Save(); 
        }

        private void Query_Load(object sender, EventArgs e)
        {
            txtQuery.Text = strQueryText;
            IsModified = false;
            IsSaved = false;
            txtQuery.SelectionLength = 0;
        }

        private void Query_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (IsModified)
            {
                DialogResult res = MessageBox.Show("Query has been modified. Would you like to save it ?",
                                 "Saving", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.Yes)
                {
                    //if (!IsQueryValid())
                    //{
                    //    MessageBox.Show("Query must have RecipientAddress field.", "Validation !", 
                    //                    MessageBoxButtons.OK,MessageBoxIcon.Warning);
                    //    e.Cancel = true;
                    //    return;
                    //}
                    Save();
                }
            }
        }
    }
}
