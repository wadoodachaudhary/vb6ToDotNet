namespace AutoNotice
{
    partial class AutoEvents
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AutoEvents));
            this.cmdSave = new System.Windows.Forms.Button();
            this.cmdClose = new System.Windows.Forms.Button();
            this.flex = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.sqlSelectCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlConn = new System.Data.SqlClient.SqlConnection();
            this.sqlDAEventsList = new System.Data.SqlClient.SqlDataAdapter();
            this.sqlCommand1 = new System.Data.SqlClient.SqlCommand();
            this.statusBar1 = new System.Windows.Forms.StatusBar();
            this.dsAutoEvents = new AutoNotice.dsAutoEvents();
            ((System.ComponentModel.ISupportInitialize)(this.flex)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dsAutoEvents)).BeginInit();
            this.SuspendLayout();
            // 
            // cmdSave
            // 
            this.cmdSave.Image = ((System.Drawing.Image)(resources.GetObject("cmdSave.Image")));
            this.cmdSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.cmdSave.Location = new System.Drawing.Point(16, 0);
            this.cmdSave.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cmdSave.Name = "cmdSave";
            this.cmdSave.Size = new System.Drawing.Size(85, 68);
            this.cmdSave.TabIndex = 10;
            this.cmdSave.Text = "Save";
            this.cmdSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.cmdSave.UseVisualStyleBackColor = true;
            this.cmdSave.Click += new System.EventHandler(this.cmdSave_Click);
            // 
            // cmdClose
            // 
            this.cmdClose.Image = ((System.Drawing.Image)(resources.GetObject("cmdClose.Image")));
            this.cmdClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.cmdClose.Location = new System.Drawing.Point(109, 0);
            this.cmdClose.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cmdClose.Name = "cmdClose";
            this.cmdClose.Size = new System.Drawing.Size(85, 68);
            this.cmdClose.TabIndex = 9;
            this.cmdClose.Text = "Exit";
            this.cmdClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.cmdClose.UseVisualStyleBackColor = true;
            this.cmdClose.Click += new System.EventHandler(this.cmdClose_Click);
            // 
            // flex
            // 
            this.flex.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.flex.ColumnInfo = resources.GetString("flex.ColumnInfo");
            this.flex.Location = new System.Drawing.Point(16, 75);
            this.flex.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.flex.Name = "flex";
            this.flex.Rows.DefaultSize = 17;
            this.flex.Size = new System.Drawing.Size(1114, 498);
            this.flex.TabIndex = 1;
            this.flex.BeforeDoubleClick += new C1.Win.C1FlexGrid.BeforeMouseDownEventHandler(this.flex_BeforeDoubleClick);
            this.flex.AfterEdit += new C1.Win.C1FlexGrid.RowColEventHandler(this.flex_AfterEdit);
            this.flex.Click += new System.EventHandler(this.flex_Click);
            this.flex.DoubleClick += new System.EventHandler(this.flex_DoubleClick);
            this.flex.KeyDown += new System.Windows.Forms.KeyEventHandler(this.flex_KeyDown);
            // 
            // sqlSelectCommand1
            // 
            this.sqlSelectCommand1.CommandText = "dbo.AutoNotice_GetAutoEventsList";
            this.sqlSelectCommand1.CommandType = System.Data.CommandType.StoredProcedure;
            this.sqlSelectCommand1.Connection = this.sqlConn;
            this.sqlSelectCommand1.Parameters.AddRange(new System.Data.SqlClient.SqlParameter[] {
            new System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, false, ((byte)(0)), ((byte)(0)), "", System.Data.DataRowVersion.Current, null)});
            // 
            // sqlConn
            // 
            this.sqlConn.ConnectionString = "Data Source=MICHAEL\\DEFAULT2000;Initial Catalog=HOMEFRONTSQLRESELLER;Integrated S" +
    "ecurity=True";
            this.sqlConn.FireInfoMessageEventOnUserErrors = false;
            // 
            // sqlDAEventsList
            // 
            this.sqlDAEventsList.SelectCommand = this.sqlSelectCommand1;
            this.sqlDAEventsList.TableMappings.AddRange(new System.Data.Common.DataTableMapping[] {
            new System.Data.Common.DataTableMapping("Table", "AutoNotice_GetAutoEventsList", new System.Data.Common.DataColumnMapping[] {
                        new System.Data.Common.DataColumnMapping("TriggerName", "TriggerName"),
                        new System.Data.Common.DataColumnMapping("Subject", "Subject"),
                        new System.Data.Common.DataColumnMapping("Message", "Message"),
                        new System.Data.Common.DataColumnMapping("StaticRecipients", "StaticRecipients"),
                        new System.Data.Common.DataColumnMapping("Buyer", "Buyer"),
                        new System.Data.Common.DataColumnMapping("Cobyer", "Cobyer"),
                        new System.Data.Common.DataColumnMapping("SalesPerson", "SalesPerson"),
                        new System.Data.Common.DataColumnMapping("PM", "PM"),
                        new System.Data.Common.DataColumnMapping("SalesManager", "SalesManager"),
                        new System.Data.Common.DataColumnMapping("Estimator", "Estimator"),
                        new System.Data.Common.DataColumnMapping("VendorPurchasing", "VendorPurchasing"),
                        new System.Data.Common.DataColumnMapping("VendorScheduling", "VendorScheduling"),
                        new System.Data.Common.DataColumnMapping("VendorService", "VendorService"),
                        new System.Data.Common.DataColumnMapping("Purchaser", "Purchaser"),
                        new System.Data.Common.DataColumnMapping("DCSalesPerson", "DCSalesPerson"),
                        new System.Data.Common.DataColumnMapping("AccountingAdministrator", "AccountingAdministrator"),
                        new System.Data.Common.DataColumnMapping("SystemAdministrator", "SystemAdministrator"),
                        new System.Data.Common.DataColumnMapping("Lender", "Lender"),
                        new System.Data.Common.DataColumnMapping("TitleCompany", "TitleCompany"),
                        new System.Data.Common.DataColumnMapping("Realtor", "Realtor"),
                        new System.Data.Common.DataColumnMapping("Broker", "Broker"),
                        new System.Data.Common.DataColumnMapping("SellingRealtor", "SellingRealtor"),
                        new System.Data.Common.DataColumnMapping("Attachment", "Attachment"),
                        new System.Data.Common.DataColumnMapping("Report", "Report"),
                        new System.Data.Common.DataColumnMapping("Param1", "Param1"),
                        new System.Data.Common.DataColumnMapping("Param2", "Param2"),
                        new System.Data.Common.DataColumnMapping("Param3", "Param3"),
                        new System.Data.Common.DataColumnMapping("Param4", "Param4"),
                        new System.Data.Common.DataColumnMapping("Param5", "Param5"),
                        new System.Data.Common.DataColumnMapping("Param6", "Param6"),
                        new System.Data.Common.DataColumnMapping("Param7", "Param7"),
                        new System.Data.Common.DataColumnMapping("Param8", "Param8"),
                        new System.Data.Common.DataColumnMapping("Param9", "Param9"),
                        new System.Data.Common.DataColumnMapping("Param10", "Param10"),
                        new System.Data.Common.DataColumnMapping("StaticCCs", "StaticCCs")})});
            this.sqlDAEventsList.UpdateCommand = this.sqlCommand1;
            // 
            // sqlCommand1
            // 
            this.sqlCommand1.CommandText = "AutoNotice_UpdateAutoEventsList";
            this.sqlCommand1.CommandType = System.Data.CommandType.StoredProcedure;
            this.sqlCommand1.Connection = this.sqlConn;
            this.sqlCommand1.Parameters.AddRange(new System.Data.SqlClient.SqlParameter[] {
            new System.Data.SqlClient.SqlParameter("@TriggerName", System.Data.SqlDbType.VarChar, 0, "TriggerName"),
            new System.Data.SqlClient.SqlParameter("@IsDisabled", System.Data.SqlDbType.Bit, 0, "IsDisabled"),
            new System.Data.SqlClient.SqlParameter("@Subject", System.Data.SqlDbType.VarChar, 0, "Subject"),
            new System.Data.SqlClient.SqlParameter("@Message", System.Data.SqlDbType.VarChar, 0, "Message"),
            new System.Data.SqlClient.SqlParameter("@StaticRecipients", System.Data.SqlDbType.VarChar, 0, "StaticRecipients"),
            new System.Data.SqlClient.SqlParameter("@StaticCCs", System.Data.SqlDbType.VarChar, 0, "StaticCCs"),
            new System.Data.SqlClient.SqlParameter("@Buyer", System.Data.SqlDbType.Bit, 0, "Buyer"),
            new System.Data.SqlClient.SqlParameter("@CoBuyer", System.Data.SqlDbType.Bit, 0, "CoBuyer"),
            new System.Data.SqlClient.SqlParameter("@SalesPerson", System.Data.SqlDbType.Bit, 0, "SalesPerson"),
            new System.Data.SqlClient.SqlParameter("@PM", System.Data.SqlDbType.Bit, 0, "PM"),
            new System.Data.SqlClient.SqlParameter("@SalesManager", System.Data.SqlDbType.Bit, 0, "SalesManager"),
            new System.Data.SqlClient.SqlParameter("@Estimator", System.Data.SqlDbType.Bit, 0, "Estimator"),
            new System.Data.SqlClient.SqlParameter("@VendorPurchasing", System.Data.SqlDbType.Bit, 0, "VendorPurchasing"),
            new System.Data.SqlClient.SqlParameter("@VendorScheduling", System.Data.SqlDbType.Bit, 0, "VendorScheduling"),
            new System.Data.SqlClient.SqlParameter("@VendorService", System.Data.SqlDbType.Bit, 0, "VendorService"),
            new System.Data.SqlClient.SqlParameter("@Purchaser", System.Data.SqlDbType.Bit, 0, "Purchaser"),
            new System.Data.SqlClient.SqlParameter("@DCSalesPerson", System.Data.SqlDbType.Bit, 0, "DCSalesPerson"),
            new System.Data.SqlClient.SqlParameter("@AccountingAdministrator", System.Data.SqlDbType.Bit, 0, "AccountingAdministrator"),
            new System.Data.SqlClient.SqlParameter("@SystemAdministrator", System.Data.SqlDbType.Bit, 0, "SystemAdministrator"),
            new System.Data.SqlClient.SqlParameter("@Lender", System.Data.SqlDbType.Bit, 0, "Lender"),
            new System.Data.SqlClient.SqlParameter("@TitleCompany", System.Data.SqlDbType.Bit, 0, "TitleCompany"),
            new System.Data.SqlClient.SqlParameter("@Realtor", System.Data.SqlDbType.Bit, 0, "Realtor"),
            new System.Data.SqlClient.SqlParameter("@Broker", System.Data.SqlDbType.Bit, 0, "Broker"),
            new System.Data.SqlClient.SqlParameter("@SellingRealtor", System.Data.SqlDbType.Bit, 0, "SellingRealtor"),
            new System.Data.SqlClient.SqlParameter("@Attachment", System.Data.SqlDbType.VarChar, 0, "Attachment"),
            new System.Data.SqlClient.SqlParameter("@Report", System.Data.SqlDbType.VarChar, 0, "Report"),
            new System.Data.SqlClient.SqlParameter("@Param1", System.Data.SqlDbType.VarChar, 0, "Param1"),
            new System.Data.SqlClient.SqlParameter("@Param2", System.Data.SqlDbType.VarChar, 0, "Param2"),
            new System.Data.SqlClient.SqlParameter("@Param3", System.Data.SqlDbType.VarChar, 0, "Param3"),
            new System.Data.SqlClient.SqlParameter("@Param4", System.Data.SqlDbType.VarChar, 0, "Param4"),
            new System.Data.SqlClient.SqlParameter("@Param5", System.Data.SqlDbType.VarChar, 0, "Param5"),
            new System.Data.SqlClient.SqlParameter("@Param6", System.Data.SqlDbType.VarChar, 0, "Param6"),
            new System.Data.SqlClient.SqlParameter("@Param7", System.Data.SqlDbType.VarChar, 0, "Param7"),
            new System.Data.SqlClient.SqlParameter("@Param8", System.Data.SqlDbType.VarChar, 0, "Param8"),
            new System.Data.SqlClient.SqlParameter("@Param9", System.Data.SqlDbType.VarChar, 0, "Param9"),
            new System.Data.SqlClient.SqlParameter("@Param10", System.Data.SqlDbType.VarChar, 0, "Param10"),
            new System.Data.SqlClient.SqlParameter("@FormattedSubject", System.Data.SqlDbType.VarChar, 0, "FormattedSubject"),
            new System.Data.SqlClient.SqlParameter("@FormattedMessage", System.Data.SqlDbType.VarChar, 0, "FormattedMessage"),
            new System.Data.SqlClient.SqlParameter("@FromAddress", System.Data.SqlDbType.VarChar, 0, "FromAddress")});
            // 
            // statusBar1
            // 
            this.statusBar1.Location = new System.Drawing.Point(0, 581);
            this.statusBar1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.statusBar1.Name = "statusBar1";
            this.statusBar1.Size = new System.Drawing.Size(1148, 27);
            this.statusBar1.TabIndex = 13;
            this.statusBar1.Text = "Use Double click to edit an event";
            // 
            // dsAutoEvents
            // 
            this.dsAutoEvents.DataSetName = "dsAutoEvents";
            this.dsAutoEvents.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // AutoEvents
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1148, 608);
            this.Controls.Add(this.statusBar1);
            this.Controls.Add(this.flex);
            this.Controls.Add(this.cmdSave);
            this.Controls.Add(this.cmdClose);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "AutoEvents";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Auto Events Editor";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.AutoEvents_FormClosing);
            this.Load += new System.EventHandler(this.AutoEvents_Load);
            ((System.ComponentModel.ISupportInitialize)(this.flex)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dsAutoEvents)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button cmdSave;
        private System.Windows.Forms.Button cmdClose;
        private C1.Win.C1FlexGrid.C1FlexGrid flex;
        private System.Data.SqlClient.SqlCommand sqlSelectCommand1;
        private System.Data.SqlClient.SqlConnection sqlConn;
        private System.Data.SqlClient.SqlDataAdapter sqlDAEventsList;
        private dsAutoEvents dsAutoEvents;
        private System.Data.SqlClient.SqlCommand sqlCommand1;
        private System.Windows.Forms.StatusBar statusBar1;
    }
}