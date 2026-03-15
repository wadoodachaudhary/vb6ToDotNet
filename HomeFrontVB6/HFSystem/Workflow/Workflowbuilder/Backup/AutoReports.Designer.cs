namespace AutoNotice
{
    partial class AutoReports
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AutoReports));
            this.flex = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.sqlSelectCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlConn = new System.Data.SqlClient.SqlConnection();
            this.sqlDAAutoReports = new System.Data.SqlClient.SqlDataAdapter();
            this.sqlCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlCommand3 = new System.Data.SqlClient.SqlCommand();
            this.sqlCommand2 = new System.Data.SqlClient.SqlCommand();
            this.cmdDelete = new System.Windows.Forms.Button();
            this.cmdSave = new System.Windows.Forms.Button();
            this.cmdClose = new System.Windows.Forms.Button();
            this.cmdNew = new System.Windows.Forms.Button();
            this.statusBar1 = new System.Windows.Forms.StatusBar();
            this.dsAutoReports = new AutoNotice.dsAutoReports();
            this.dsAutoReportsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.flex)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dsAutoReports)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dsAutoReportsBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // flex
            // 
            this.flex.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.flex.ColumnInfo = resources.GetString("flex.ColumnInfo");
            this.flex.Location = new System.Drawing.Point(1, 59);
            this.flex.Name = "flex";
            this.flex.Rows.Count = 2;
            this.flex.Rows.DefaultSize = 17;
            this.flex.SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.RowRange;
            this.flex.Size = new System.Drawing.Size(954, 407);
            this.flex.TabIndex = 1;
            this.flex.KeyDown += new System.Windows.Forms.KeyEventHandler(this.flex_KeyDown);
            this.flex.DoubleClick += new System.EventHandler(this.flex_DoubleClick);
            this.flex.BeforeDoubleClick += new C1.Win.C1FlexGrid.BeforeMouseDownEventHandler(this.flex_BeforeDoubleClick);
            this.flex.AfterEdit += new C1.Win.C1FlexGrid.RowColEventHandler(this.flex_AfterEdit);
            // 
            // sqlSelectCommand1
            // 
            this.sqlSelectCommand1.CommandText = "dbo.AutoNotice_GetAutoReportsList";
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
            // sqlDAAutoReports
            // 
            this.sqlDAAutoReports.DeleteCommand = this.sqlCommand1;
            this.sqlDAAutoReports.InsertCommand = this.sqlCommand3;
            this.sqlDAAutoReports.SelectCommand = this.sqlSelectCommand1;
            this.sqlDAAutoReports.TableMappings.AddRange(new System.Data.Common.DataTableMapping[] {
            new System.Data.Common.DataTableMapping("Table", "AutoNotice_GetAutoReportsList", new System.Data.Common.DataColumnMapping[] {
                        new System.Data.Common.DataColumnMapping("Seq", "Seq"),
                        new System.Data.Common.DataColumnMapping("RecipientAddress", "RecipientAddress"),
                        new System.Data.Common.DataColumnMapping("FromAddress", "FromAddress"),
                        new System.Data.Common.DataColumnMapping("Subject", "Subject"),
                        new System.Data.Common.DataColumnMapping("Message", "Message"),
                        new System.Data.Common.DataColumnMapping("FrequencyType", "FrequencyType"),
                        new System.Data.Common.DataColumnMapping("FrequencyValue", "FrequencyValue"),
                        new System.Data.Common.DataColumnMapping("TimeofDay", "TimeofDay"),
                        new System.Data.Common.DataColumnMapping("NextRun", "NextRun"),
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
                        new System.Data.Common.DataColumnMapping("Param10", "Param10")})});
            this.sqlDAAutoReports.UpdateCommand = this.sqlCommand2;
            // 
            // sqlCommand1
            // 
            this.sqlCommand1.CommandText = "AutoNotice_DeleteAutoReportsItem";
            this.sqlCommand1.CommandType = System.Data.CommandType.StoredProcedure;
            this.sqlCommand1.Connection = this.sqlConn;
            this.sqlCommand1.Parameters.AddRange(new System.Data.SqlClient.SqlParameter[] {
            new System.Data.SqlClient.SqlParameter("@Seq", System.Data.SqlDbType.Int, 0, "Seq")});
            // 
            // sqlCommand3
            // 
            this.sqlCommand3.CommandText = "AutoNotice_InsertAutoReportItem";
            this.sqlCommand3.CommandType = System.Data.CommandType.StoredProcedure;
            this.sqlCommand3.Connection = this.sqlConn;
            this.sqlCommand3.Parameters.AddRange(new System.Data.SqlClient.SqlParameter[] {
            new System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 50, "Description"),
            new System.Data.SqlClient.SqlParameter("@RecipientAddress", System.Data.SqlDbType.VarChar, 0, "RecipientAddress"),
            new System.Data.SqlClient.SqlParameter("@RecipientQuery", System.Data.SqlDbType.VarChar, 0, "RecipientQuery"),
            new System.Data.SqlClient.SqlParameter("@FromAddress", System.Data.SqlDbType.VarChar, 0, "FromAddress"),
            new System.Data.SqlClient.SqlParameter("@Subject", System.Data.SqlDbType.VarChar, 0, "Subject"),
            new System.Data.SqlClient.SqlParameter("@Message", System.Data.SqlDbType.VarChar, 0, "Message"),
            new System.Data.SqlClient.SqlParameter("@FrequencyType", System.Data.SqlDbType.VarChar, 0, "FrequencyType"),
            new System.Data.SqlClient.SqlParameter("@FrequencyValue", System.Data.SqlDbType.Int, 0, "FrequencyValue"),
            new System.Data.SqlClient.SqlParameter("@TimeofDay", System.Data.SqlDbType.DateTime, 0, "TimeofDay"),
            new System.Data.SqlClient.SqlParameter("@NextRun", System.Data.SqlDbType.DateTime, 0, "NextRun"),
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
            new System.Data.SqlClient.SqlParameter("@FormattedMessage", System.Data.SqlDbType.VarChar, 0, "FormattedMessage")});
            // 
            // sqlCommand2
            // 
            this.sqlCommand2.CommandText = "AutoNotice_UpdateAutoReportsList";
            this.sqlCommand2.CommandType = System.Data.CommandType.StoredProcedure;
            this.sqlCommand2.Connection = this.sqlConn;
            this.sqlCommand2.Parameters.AddRange(new System.Data.SqlClient.SqlParameter[] {
            new System.Data.SqlClient.SqlParameter("@Seq", System.Data.SqlDbType.Int, 0, "Seq"),
            new System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 50, "Description"),
            new System.Data.SqlClient.SqlParameter("@RecipientAddress", System.Data.SqlDbType.VarChar, 0, "RecipientAddress"),
            new System.Data.SqlClient.SqlParameter("@FromAddress", System.Data.SqlDbType.VarChar, 0, "FromAddress"),
            new System.Data.SqlClient.SqlParameter("@RecipientQuery", System.Data.SqlDbType.VarChar, 0, "RecipientQuery"),
            new System.Data.SqlClient.SqlParameter("@Subject", System.Data.SqlDbType.VarChar, 0, "Subject"),
            new System.Data.SqlClient.SqlParameter("@Message", System.Data.SqlDbType.VarChar, 0, "Message"),
            new System.Data.SqlClient.SqlParameter("@FrequencyType", System.Data.SqlDbType.VarChar, 0, "FrequencyType"),
            new System.Data.SqlClient.SqlParameter("@FrequencyValue", System.Data.SqlDbType.Int, 0, "FrequencyValue"),
            new System.Data.SqlClient.SqlParameter("@TimeofDay", System.Data.SqlDbType.DateTime, 0, "TimeofDay"),
            new System.Data.SqlClient.SqlParameter("@NextRun", System.Data.SqlDbType.DateTime, 0, "NextRun"),
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
            new System.Data.SqlClient.SqlParameter("@FormattedSubject", System.Data.SqlDbType.VarChar, 2000, "FormattedSubject"),
            new System.Data.SqlClient.SqlParameter("@FormattedMessage", System.Data.SqlDbType.VarChar, 8000, "FormattedMessage")});
            // 
            // cmdDelete
            // 
            this.cmdDelete.Image = ((System.Drawing.Image)(resources.GetObject("cmdDelete.Image")));
            this.cmdDelete.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.cmdDelete.Location = new System.Drawing.Point(122, 0);
            this.cmdDelete.Name = "cmdDelete";
            this.cmdDelete.Size = new System.Drawing.Size(64, 55);
            this.cmdDelete.TabIndex = 7;
            this.cmdDelete.Text = "Delete";
            this.cmdDelete.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.cmdDelete.UseVisualStyleBackColor = true;
            this.cmdDelete.Click += new System.EventHandler(this.cmdDelete_Click);
            // 
            // cmdSave
            // 
            this.cmdSave.Image = ((System.Drawing.Image)(resources.GetObject("cmdSave.Image")));
            this.cmdSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.cmdSave.Location = new System.Drawing.Point(61, 0);
            this.cmdSave.Name = "cmdSave";
            this.cmdSave.Size = new System.Drawing.Size(64, 55);
            this.cmdSave.TabIndex = 6;
            this.cmdSave.Text = "Save";
            this.cmdSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.cmdSave.UseVisualStyleBackColor = true;
            this.cmdSave.Click += new System.EventHandler(this.cmdSave_Click);
            // 
            // cmdClose
            // 
            this.cmdClose.Image = ((System.Drawing.Image)(resources.GetObject("cmdClose.Image")));
            this.cmdClose.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.cmdClose.Location = new System.Drawing.Point(203, 0);
            this.cmdClose.Name = "cmdClose";
            this.cmdClose.Size = new System.Drawing.Size(64, 55);
            this.cmdClose.TabIndex = 5;
            this.cmdClose.Text = "Exit";
            this.cmdClose.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.cmdClose.UseVisualStyleBackColor = true;
            this.cmdClose.Click += new System.EventHandler(this.cmdClose_Click);
            // 
            // cmdNew
            // 
            this.cmdNew.Image = ((System.Drawing.Image)(resources.GetObject("cmdNew.Image")));
            this.cmdNew.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.cmdNew.Location = new System.Drawing.Point(0, 0);
            this.cmdNew.Name = "cmdNew";
            this.cmdNew.Size = new System.Drawing.Size(64, 55);
            this.cmdNew.TabIndex = 4;
            this.cmdNew.Text = "New";
            this.cmdNew.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.cmdNew.UseVisualStyleBackColor = true;
            this.cmdNew.Click += new System.EventHandler(this.cmdNew_Click);
            // 
            // statusBar1
            // 
            this.statusBar1.Location = new System.Drawing.Point(0, 472);
            this.statusBar1.Name = "statusBar1";
            this.statusBar1.Size = new System.Drawing.Size(959, 22);
            this.statusBar1.TabIndex = 9;
            this.statusBar1.Text = "Use Double click to edit a report";
            // 
            // dsAutoReports
            // 
            this.dsAutoReports.DataSetName = "dsAutoReports";
            this.dsAutoReports.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // dsAutoReportsBindingSource
            // 
            this.dsAutoReportsBindingSource.DataSource = this.dsAutoReports;
            this.dsAutoReportsBindingSource.Position = 0;
            // 
            // AutoReports
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(959, 494);
            this.Controls.Add(this.statusBar1);
            this.Controls.Add(this.flex);
            this.Controls.Add(this.cmdDelete);
            this.Controls.Add(this.cmdSave);
            this.Controls.Add(this.cmdClose);
            this.Controls.Add(this.cmdNew);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "AutoReports";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Auto Reports";
            this.Load += new System.EventHandler(this.AutoReports_Load);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.AutoReports_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.flex)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dsAutoReports)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dsAutoReportsBindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button cmdNew;
        private System.Windows.Forms.Button cmdDelete;
        private System.Windows.Forms.Button cmdSave;
        private System.Windows.Forms.Button cmdClose;
        private C1.Win.C1FlexGrid.C1FlexGrid flex;
        private System.Data.SqlClient.SqlCommand sqlSelectCommand1;
        private System.Data.SqlClient.SqlDataAdapter sqlDAAutoReports;
        private System.Data.SqlClient.SqlConnection sqlConn;
        private dsAutoReports dsAutoReports;
        private System.Windows.Forms.BindingSource dsAutoReportsBindingSource;
        private System.Data.SqlClient.SqlCommand sqlCommand1;
        private System.Data.SqlClient.SqlCommand sqlCommand2;
        private System.Data.SqlClient.SqlCommand sqlCommand3;
        private System.Windows.Forms.StatusBar statusBar1;

    }
}