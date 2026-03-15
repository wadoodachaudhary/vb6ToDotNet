namespace AutoNotice
{
    partial class TriggerParamDetails
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
            this.lblTable = new System.Windows.Forms.Label();
            this.grvTable = new System.Windows.Forms.DataGridView();
            this.FieldName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FieldType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dsTriggerParamsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.dsTriggerParams = new AutoNotice.dsTriggerParams();
            this.lblParamValue = new System.Windows.Forms.Label();
            this.txtValue = new System.Windows.Forms.TextBox();
            this.cmdOK = new System.Windows.Forms.Button();
            this.cmdCancel = new System.Windows.Forms.Button();
            this.sqlSelectCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlConn = new System.Data.SqlClient.SqlConnection();
            this.sqlDATrigParams = new System.Data.SqlClient.SqlDataAdapter();
            this.lblSearch = new System.Windows.Forms.Label();
            this.txtParamName = new System.Windows.Forms.TextBox();
            this.cmdSearch = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.grvTable)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dsTriggerParamsBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dsTriggerParams)).BeginInit();
            this.SuspendLayout();
            // 
            // lblTable
            // 
            this.lblTable.AutoSize = true;
            this.lblTable.Location = new System.Drawing.Point(12, 9);
            this.lblTable.Name = "lblTable";
            this.lblTable.Size = new System.Drawing.Size(178, 13);
            this.lblTable.TabIndex = 0;
            this.lblTable.Text = "Select parameter value from the list !";
            // 
            // grvTable
            // 
            this.grvTable.AllowUserToAddRows = false;
            this.grvTable.AllowUserToDeleteRows = false;
            this.grvTable.AutoGenerateColumns = false;
            this.grvTable.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grvTable.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.FieldName,
            this.FieldType});
            this.grvTable.DataMember = "AutoNotice_GetTriggersParamsList";
            this.grvTable.DataSource = this.dsTriggerParamsBindingSource;
            this.grvTable.Location = new System.Drawing.Point(15, 71);
            this.grvTable.Name = "grvTable";
            this.grvTable.ReadOnly = true;
            this.grvTable.Size = new System.Drawing.Size(371, 245);
            this.grvTable.TabIndex = 2;
            this.grvTable.DoubleClick += new System.EventHandler(this.grvTable_DoubleClick);
            this.grvTable.KeyDown += new System.Windows.Forms.KeyEventHandler(this.grvTable_KeyDown);
            // 
            // FieldName
            // 
            this.FieldName.DataPropertyName = "FieldName";
            this.FieldName.HeaderText = "Parameter";
            this.FieldName.Name = "FieldName";
            this.FieldName.ReadOnly = true;
            this.FieldName.Width = 250;
            // 
            // FieldType
            // 
            this.FieldType.DataPropertyName = "FieldType";
            this.FieldType.HeaderText = "Type";
            this.FieldType.Name = "FieldType";
            this.FieldType.ReadOnly = true;
            this.FieldType.Visible = false;
            // 
            // dsTriggerParamsBindingSource
            // 
            this.dsTriggerParamsBindingSource.DataSource = this.dsTriggerParams;
            this.dsTriggerParamsBindingSource.Position = 0;
            // 
            // dsTriggerParams
            // 
            this.dsTriggerParams.DataSetName = "dsTriggerParams";
            this.dsTriggerParams.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // lblParamValue
            // 
            this.lblParamValue.AutoSize = true;
            this.lblParamValue.Location = new System.Drawing.Point(14, 325);
            this.lblParamValue.Name = "lblParamValue";
            this.lblParamValue.Size = new System.Drawing.Size(84, 13);
            this.lblParamValue.TabIndex = 4;
            this.lblParamValue.Text = "Parameter value";
            this.lblParamValue.Visible = false;
            // 
            // txtValue
            // 
            this.txtValue.Location = new System.Drawing.Point(98, 322);
            this.txtValue.Name = "txtValue";
            this.txtValue.Size = new System.Drawing.Size(290, 20);
            this.txtValue.TabIndex = 5;
            this.txtValue.Visible = false;
            // 
            // cmdOK
            // 
            this.cmdOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.cmdOK.Location = new System.Drawing.Point(234, 349);
            this.cmdOK.Name = "cmdOK";
            this.cmdOK.Size = new System.Drawing.Size(75, 23);
            this.cmdOK.TabIndex = 6;
            this.cmdOK.Text = "OK";
            this.cmdOK.UseVisualStyleBackColor = true;
            this.cmdOK.Click += new System.EventHandler(this.cmdOK_Click);
            // 
            // cmdCancel
            // 
            this.cmdCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.cmdCancel.Location = new System.Drawing.Point(315, 349);
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.Size = new System.Drawing.Size(75, 23);
            this.cmdCancel.TabIndex = 7;
            this.cmdCancel.Text = "Cancel";
            this.cmdCancel.UseVisualStyleBackColor = true;
            this.cmdCancel.Click += new System.EventHandler(this.cmdCancel_Click);
            // 
            // sqlSelectCommand1
            // 
            this.sqlSelectCommand1.CommandText = "dbo.AutoNotice_GetTriggersParamsList";
            this.sqlSelectCommand1.CommandType = System.Data.CommandType.StoredProcedure;
            this.sqlSelectCommand1.Connection = this.sqlConn;
            this.sqlSelectCommand1.Parameters.AddRange(new System.Data.SqlClient.SqlParameter[] {
            new System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, false, ((byte)(0)), ((byte)(0)), "", System.Data.DataRowVersion.Current, null),
            new System.Data.SqlClient.SqlParameter("@TableName", System.Data.SqlDbType.VarChar, 100),
            new System.Data.SqlClient.SqlParameter("@TriggerName", System.Data.SqlDbType.VarChar, 100)});
            // 
            // sqlConn
            // 
            this.sqlConn.ConnectionString = "Data Source=MICHAEL\\SQL2005;Initial Catalog=HomeFront;Integrated Security=True";
            this.sqlConn.FireInfoMessageEventOnUserErrors = false;
            // 
            // sqlDATrigParams
            // 
            this.sqlDATrigParams.SelectCommand = this.sqlSelectCommand1;
            this.sqlDATrigParams.TableMappings.AddRange(new System.Data.Common.DataTableMapping[] {
            new System.Data.Common.DataTableMapping("Table", "AutoNotice_GetTriggersParamsList", new System.Data.Common.DataColumnMapping[] {
                        new System.Data.Common.DataColumnMapping("FieldName", "FieldName")})});
            // 
            // lblSearch
            // 
            this.lblSearch.AutoSize = true;
            this.lblSearch.Location = new System.Drawing.Point(12, 35);
            this.lblSearch.Name = "lblSearch";
            this.lblSearch.Size = new System.Drawing.Size(86, 13);
            this.lblSearch.TabIndex = 8;
            this.lblSearch.Text = "Parameter Name";
            // 
            // txtParamName
            // 
            this.txtParamName.Location = new System.Drawing.Point(104, 32);
            this.txtParamName.Name = "txtParamName";
            this.txtParamName.Size = new System.Drawing.Size(201, 20);
            this.txtParamName.TabIndex = 9;
            this.txtParamName.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtParamName_KeyDown);
            // 
            // cmdSearch
            // 
            this.cmdSearch.Location = new System.Drawing.Point(311, 30);
            this.cmdSearch.Name = "cmdSearch";
            this.cmdSearch.Size = new System.Drawing.Size(75, 23);
            this.cmdSearch.TabIndex = 10;
            this.cmdSearch.Text = "Search";
            this.cmdSearch.UseVisualStyleBackColor = true;
            this.cmdSearch.Click += new System.EventHandler(this.cmdSearch_Click);
            // 
            // TriggerParamDetails
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(400, 384);
            this.Controls.Add(this.cmdSearch);
            this.Controls.Add(this.txtParamName);
            this.Controls.Add(this.lblSearch);
            this.Controls.Add(this.cmdCancel);
            this.Controls.Add(this.cmdOK);
            this.Controls.Add(this.txtValue);
            this.Controls.Add(this.lblParamValue);
            this.Controls.Add(this.grvTable);
            this.Controls.Add(this.lblTable);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "TriggerParamDetails";
            this.Text = "Select Value";
            this.Load += new System.EventHandler(this.TriggerParamDetails_Load);
            ((System.ComponentModel.ISupportInitialize)(this.grvTable)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dsTriggerParamsBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dsTriggerParams)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblTable;
        private System.Windows.Forms.DataGridView grvTable;
        private System.Windows.Forms.Label lblParamValue;
        private System.Windows.Forms.TextBox txtValue;
        private System.Windows.Forms.Button cmdOK;
        private System.Windows.Forms.Button cmdCancel;
        private System.Data.SqlClient.SqlCommand sqlSelectCommand1;
        private System.Data.SqlClient.SqlConnection sqlConn;
        private System.Data.SqlClient.SqlDataAdapter sqlDATrigParams;
        private System.Windows.Forms.BindingSource dsTriggerParamsBindingSource;
        private dsTriggerParams dsTriggerParams;
        private System.Windows.Forms.Label lblSearch;
        private System.Windows.Forms.TextBox txtParamName;
        private System.Windows.Forms.Button cmdSearch;
        private System.Windows.Forms.DataGridViewTextBoxColumn FieldName;
        private System.Windows.Forms.DataGridViewTextBoxColumn FieldType;
    }
}