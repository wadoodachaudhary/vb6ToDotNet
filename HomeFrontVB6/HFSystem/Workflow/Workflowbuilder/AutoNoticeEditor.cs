using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace AutoNotice
{
    public partial class AutoNoticeEditor : Form
    {
        public AutoNoticeEditor()
        {
            InitializeComponent();
        }

        /// <summary>
        /// HFSystem.dll
        /// </summary>
        HFSystem.Application hfsys;

        /// <summary>
        /// The connection string to the sql data adapter
        /// </summary>
        private string sqlConnectionString;

        /// <summary>
        /// The ODBC DSN name
        /// </summary>
        private string DataSourceName;

        /// <summary>
        /// DB User ID
        /// </summary>
        private string DB_UID;

        /// <summary>
        /// DB Password
        /// </summary>
        private string DB_PWD;

        /// <summary>
        /// Current database name
        /// </summary>
        private string DatabaseName;

        /// <summary>
        /// The application user login
        /// </summary>
        private string appLogin;



        /// <summary>
        /// Returns the database server name by DSN name
        /// </summary>
        /// <param name="userDSN"></param>
        /// <returns></returns>
        private string GetDBServerName(string userDSN)
        {
            // get reference to the HKLM registry key...
            RegistryKey rkHKLM = Registry.LocalMachine;
            RegistryKey rkRun;
            string ServerName;

            try
            {
                rkRun = rkHKLM.OpenSubKey("Software\\Odbc\\odbc.ini\\" + userDSN);
                ServerName = rkRun.GetValue("Server").ToString();
                rkRun.Close();
                rkHKLM.Close();
                return ServerName;
            }
            catch
            {
                // error while opening the subkey...
                MessageBox.Show("Unable to open the SOFTWARE subkey!", "Getting server name");
                // close the HKLM key...
                rkHKLM.Close();
                return null;
            }
        }

        private void AutoNoticeEditor_Load(object sender, EventArgs e)
        {
            try
            {
                string version;
                if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
                    version = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
                else
                    version = Application.ProductVersion;

                string ExePath = System.AppDomain.CurrentDomain.BaseDirectory;
                ExePath = ExePath.Remove(ExePath.Length - 1);

                hfsys = new HFSystem.Application();
                //"hfsched"
                if (hfsys.Login("Workflowbuilder", "Workflow Builder", version, ExePath))
                {
                    // application login
                    appLogin = hfsys.LoginID;

                    // Data Source Name
                    DataSourceName = hfsys.LoginDSN;
                    DB_UID = hfsys.LoginDBUID;
                    DB_PWD = hfsys.LoginDBPWD;

                    if (DB_UID == "")
                    {
                        // Sql connection string
                        sqlConnectionString = "Data Source=" + GetDBServerName(DataSourceName) + ";" +
                                              "Initial Catalog=" + hfsys.LoginDBName + ";" +
                                              "Integrated Security=True;Persist Security Info=True;Connection Timeout=0";
                    }
                    else
                    {
                        // Sql connection string
                        sqlConnectionString = "Data Source=" + GetDBServerName(DataSourceName) + ";" +
                                              "Initial Catalog=" + hfsys.LoginDBName + ";" +
                                              "User Id=" + DB_UID + ";Password=" + DB_PWD + ";" +
                                              "Persist Security Info=True;Connection Timeout=0";
                    }
                }
                else
                {
                    Application.Exit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        /// <summary>
        /// Loads auto reports form
        /// </summary>
        private void OnAutoReports()
        {
            AutoReports stdReportNotes = new AutoReports();
            stdReportNotes.ConnectionString = sqlConnectionString;
            stdReportNotes.ShowDialog(this);
        }

        private void mnuExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void aboutItem_Click(object sender, EventArgs e)
        {
            hfsys.About();
        }

        private void picAutoReports_Click(object sender, EventArgs e)
        {
            OnAutoReports();
        }

        private void picAutoEvents_Click(object sender, EventArgs e)
        {
            AutoEvents trigger_table = new AutoEvents();
            trigger_table.DSN_NAME = DataSourceName;
            trigger_table.DB_USER = DB_UID;
            trigger_table.DB_PASSWORD = DB_PWD;
            trigger_table.ConnectionString = sqlConnectionString;
            trigger_table.ShowDialog(this);
        }

        private void picAutoReports_MouseHover(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Hand;
        }

        private void picAutoReports_MouseLeave(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Default;
        }

        private void picAutoEvents_MouseHover(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Hand;
        }

        private void picAutoEvents_MouseLeave(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Default;
        }

        
                        
    }
}
