namespace AutoNotice
{
    partial class AutoReportDetails
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AutoReportDetails));
            this.label1 = new System.Windows.Forms.Label();
            this.txtRecipient = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtSender = new System.Windows.Forms.TextBox();
            this.txtSubject = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtMessage = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtAttachment = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtReport = new System.Windows.Forms.TextBox();
            this.picAttachment = new System.Windows.Forms.PictureBox();
            this.picReport = new System.Windows.Forms.PictureBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.grFrequency = new System.Windows.Forms.GroupBox();
            this.cmbFrequencyValue = new System.Windows.Forms.ComboBox();
            this.cmbFrequencyType = new System.Windows.Forms.ComboBox();
            this.grRun = new System.Windows.Forms.GroupBox();
            this.nextRun = new System.Windows.Forms.DateTimePicker();
            this.label10 = new System.Windows.Forms.Label();
            this.runTime = new System.Windows.Forms.DateTimePicker();
            this.label9 = new System.Windows.Forms.Label();
            this.lblParameters = new System.Windows.Forms.Label();
            this.cmdSave = new System.Windows.Forms.Button();
            this.cmdCancel = new System.Windows.Forms.Button();
            this.openAttachment = new System.Windows.Forms.OpenFileDialog();
            this.openReport = new System.Windows.Forms.OpenFileDialog();
            this.picMessage = new System.Windows.Forms.PictureBox();
            this.cmdQuery = new System.Windows.Forms.Button();
            this.lblRecipientQuery = new System.Windows.Forms.Label();
            this.txtRecQuery = new System.Windows.Forms.TextBox();
            this.picSubject = new System.Windows.Forms.PictureBox();
            this.lblDescription = new System.Windows.Forms.Label();
            this.txtDescription = new System.Windows.Forms.TextBox();
            this.flexParams = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.webSubject = new System.Windows.Forms.WebBrowser();
            this.webBrowser = new System.Windows.Forms.WebBrowser();
            this.cmdTestMe = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.picAttachment)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picReport)).BeginInit();
            this.grFrequency.SuspendLayout();
            this.grRun.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picMessage)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picSubject)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.flexParams)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 44);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Recipient:";
            // 
            // txtRecipient
            // 
            this.txtRecipient.Location = new System.Drawing.Point(80, 41);
            this.txtRecipient.Name = "txtRecipient";
            this.txtRecipient.Size = new System.Drawing.Size(337, 20);
            this.txtRecipient.TabIndex = 1;
            this.txtRecipient.TextChanged += new System.EventHandler(this.txtRecipient_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(23, 72);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(44, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Sender:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(21, 101);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(46, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "Subject:";
            // 
            // txtSender
            // 
            this.txtSender.Location = new System.Drawing.Point(80, 69);
            this.txtSender.Name = "txtSender";
            this.txtSender.Size = new System.Drawing.Size(337, 20);
            this.txtSender.TabIndex = 4;
            this.txtSender.TextChanged += new System.EventHandler(this.txtSender_TextChanged);
            // 
            // txtSubject
            // 
            this.txtSubject.Location = new System.Drawing.Point(444, 145);
            this.txtSubject.Name = "txtSubject";
            this.txtSubject.Size = new System.Drawing.Size(79, 20);
            this.txtSubject.TabIndex = 5;
            this.txtSubject.Visible = false;
            this.txtSubject.TextChanged += new System.EventHandler(this.txtSubject_TextChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(14, 241);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Message:";
            // 
            // txtMessage
            // 
            this.txtMessage.Location = new System.Drawing.Point(444, 295);
            this.txtMessage.Multiline = true;
            this.txtMessage.Name = "txtMessage";
            this.txtMessage.Size = new System.Drawing.Size(79, 31);
            this.txtMessage.TabIndex = 7;
            this.txtMessage.Visible = false;
            this.txtMessage.TextChanged += new System.EventHandler(this.txtMessage_TextChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(459, 9);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(64, 13);
            this.label5.TabIndex = 8;
            this.label5.Text = "Attachment:";
            // 
            // txtAttachment
            // 
            this.txtAttachment.Location = new System.Drawing.Point(536, 9);
            this.txtAttachment.Name = "txtAttachment";
            this.txtAttachment.Size = new System.Drawing.Size(337, 20);
            this.txtAttachment.TabIndex = 9;
            this.txtAttachment.TextChanged += new System.EventHandler(this.txtAttachment_TextChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(484, 37);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(39, 13);
            this.label6.TabIndex = 10;
            this.label6.Text = "Report";
            // 
            // txtReport
            // 
            this.txtReport.Location = new System.Drawing.Point(536, 37);
            this.txtReport.Name = "txtReport";
            this.txtReport.Size = new System.Drawing.Size(337, 20);
            this.txtReport.TabIndex = 11;
            this.txtReport.TextChanged += new System.EventHandler(this.txtReport_TextChanged);
            // 
            // picAttachment
            // 
            this.picAttachment.Image = ((System.Drawing.Image)(resources.GetObject("picAttachment.Image")));
            this.picAttachment.Location = new System.Drawing.Point(870, 9);
            this.picAttachment.Name = "picAttachment";
            this.picAttachment.Size = new System.Drawing.Size(26, 20);
            this.picAttachment.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.picAttachment.TabIndex = 24;
            this.picAttachment.TabStop = false;
            this.picAttachment.Click += new System.EventHandler(this.picAttachment_Click);
            // 
            // picReport
            // 
            this.picReport.Image = ((System.Drawing.Image)(resources.GetObject("picReport.Image")));
            this.picReport.Location = new System.Drawing.Point(870, 37);
            this.picReport.Name = "picReport";
            this.picReport.Size = new System.Drawing.Size(26, 20);
            this.picReport.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.picReport.TabIndex = 25;
            this.picReport.TabStop = false;
            this.picReport.Click += new System.EventHandler(this.picReport_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(6, 31);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(34, 13);
            this.label7.TabIndex = 26;
            this.label7.Text = "Type:";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(3, 64);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(37, 13);
            this.label8.TabIndex = 27;
            this.label8.Text = "Value:";
            // 
            // grFrequency
            // 
            this.grFrequency.Controls.Add(this.cmbFrequencyValue);
            this.grFrequency.Controls.Add(this.cmbFrequencyType);
            this.grFrequency.Controls.Add(this.label8);
            this.grFrequency.Controls.Add(this.label7);
            this.grFrequency.Location = new System.Drawing.Point(536, 267);
            this.grFrequency.Name = "grFrequency";
            this.grFrequency.Size = new System.Drawing.Size(220, 90);
            this.grFrequency.TabIndex = 28;
            this.grFrequency.TabStop = false;
            this.grFrequency.Text = "Frequency:";
            // 
            // cmbFrequencyValue
            // 
            this.cmbFrequencyValue.FormattingEnabled = true;
            this.cmbFrequencyValue.Location = new System.Drawing.Point(55, 60);
            this.cmbFrequencyValue.Name = "cmbFrequencyValue";
            this.cmbFrequencyValue.Size = new System.Drawing.Size(141, 21);
            this.cmbFrequencyValue.TabIndex = 28;
            this.cmbFrequencyValue.SelectedValueChanged += new System.EventHandler(this.cmbFrequencyValue_SelectedValueChanged);
            this.cmbFrequencyValue.TextChanged += new System.EventHandler(this.cmbFrequencyValue_TextChanged);
            // 
            // cmbFrequencyType
            // 
            this.cmbFrequencyType.FormattingEnabled = true;
            this.cmbFrequencyType.Items.AddRange(new object[] {
            "Day of week",
            "Day of month",
            "Every nth week",
            "Every nth month"});
            this.cmbFrequencyType.Location = new System.Drawing.Point(55, 28);
            this.cmbFrequencyType.Name = "cmbFrequencyType";
            this.cmbFrequencyType.Size = new System.Drawing.Size(141, 21);
            this.cmbFrequencyType.TabIndex = 27;
            this.cmbFrequencyType.SelectedIndexChanged += new System.EventHandler(this.cmbFrequencyType_SelectedIndexChanged);
            this.cmbFrequencyType.SelectedValueChanged += new System.EventHandler(this.cmbFrequencyType_SelectedValueChanged);
            // 
            // grRun
            // 
            this.grRun.Controls.Add(this.nextRun);
            this.grRun.Controls.Add(this.label10);
            this.grRun.Controls.Add(this.runTime);
            this.grRun.Controls.Add(this.label9);
            this.grRun.Location = new System.Drawing.Point(536, 363);
            this.grRun.Name = "grRun";
            this.grRun.Size = new System.Drawing.Size(220, 90);
            this.grRun.TabIndex = 29;
            this.grRun.TabStop = false;
            this.grRun.Text = "Running";
            // 
            // nextRun
            // 
            this.nextRun.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.nextRun.Location = new System.Drawing.Point(78, 52);
            this.nextRun.Name = "nextRun";
            this.nextRun.Size = new System.Drawing.Size(106, 20);
            this.nextRun.TabIndex = 30;
            this.nextRun.ValueChanged += new System.EventHandler(this.nextRun_ValueChanged);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(7, 56);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(55, 13);
            this.label10.TabIndex = 30;
            this.label10.Text = "Next Run:";
            // 
            // runTime
            // 
            this.runTime.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.runTime.Location = new System.Drawing.Point(78, 24);
            this.runTime.Name = "runTime";
            this.runTime.ShowUpDown = true;
            this.runTime.Size = new System.Drawing.Size(106, 20);
            this.runTime.TabIndex = 30;
            this.runTime.ValueChanged += new System.EventHandler(this.runTime_ValueChanged);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(6, 28);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(56, 13);
            this.label9.TabIndex = 30;
            this.label9.Text = "Run Time:";
            // 
            // lblParameters
            // 
            this.lblParameters.AutoSize = true;
            this.lblParameters.Location = new System.Drawing.Point(463, 63);
            this.lblParameters.Name = "lblParameters";
            this.lblParameters.Size = new System.Drawing.Size(60, 13);
            this.lblParameters.TabIndex = 30;
            this.lblParameters.Text = "Parameters";
            // 
            // cmdSave
            // 
            this.cmdSave.Location = new System.Drawing.Point(736, 479);
            this.cmdSave.Name = "cmdSave";
            this.cmdSave.Size = new System.Drawing.Size(75, 23);
            this.cmdSave.TabIndex = 32;
            this.cmdSave.Text = "Save";
            this.cmdSave.UseVisualStyleBackColor = true;
            this.cmdSave.Click += new System.EventHandler(this.cmdSave_Click);
            // 
            // cmdCancel
            // 
            this.cmdCancel.Location = new System.Drawing.Point(818, 479);
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.Size = new System.Drawing.Size(75, 23);
            this.cmdCancel.TabIndex = 33;
            this.cmdCancel.Text = "Cancel";
            this.cmdCancel.UseVisualStyleBackColor = true;
            this.cmdCancel.Click += new System.EventHandler(this.cmdCancel_Click);
            // 
            // openReport
            // 
            this.openReport.Filter = "(*.rpt)|*.rpt";
            // 
            // picMessage
            // 
            this.picMessage.Image = ((System.Drawing.Image)(resources.GetObject("picMessage.Image")));
            this.picMessage.Location = new System.Drawing.Point(414, 241);
            this.picMessage.Name = "picMessage";
            this.picMessage.Size = new System.Drawing.Size(26, 20);
            this.picMessage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.picMessage.TabIndex = 39;
            this.picMessage.TabStop = false;
            this.picMessage.Click += new System.EventHandler(this.picMessage_Click);
            // 
            // cmdQuery
            // 
            this.cmdQuery.Location = new System.Drawing.Point(414, 377);
            this.cmdQuery.Name = "cmdQuery";
            this.cmdQuery.Size = new System.Drawing.Size(40, 24);
            this.cmdQuery.TabIndex = 42;
            this.cmdQuery.Text = "SQL";
            this.cmdQuery.UseVisualStyleBackColor = true;
            this.cmdQuery.Click += new System.EventHandler(this.cmdQuery_Click);
            // 
            // lblRecipientQuery
            // 
            this.lblRecipientQuery.Location = new System.Drawing.Point(7, 377);
            this.lblRecipientQuery.Name = "lblRecipientQuery";
            this.lblRecipientQuery.Size = new System.Drawing.Size(60, 33);
            this.lblRecipientQuery.TabIndex = 43;
            this.lblRecipientQuery.Text = "Recipient Query";
            // 
            // txtRecQuery
            // 
            this.txtRecQuery.BackColor = System.Drawing.SystemColors.Window;
            this.txtRecQuery.Location = new System.Drawing.Point(80, 377);
            this.txtRecQuery.Multiline = true;
            this.txtRecQuery.Name = "txtRecQuery";
            this.txtRecQuery.ReadOnly = true;
            this.txtRecQuery.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtRecQuery.Size = new System.Drawing.Size(337, 118);
            this.txtRecQuery.TabIndex = 44;
            this.txtRecQuery.TextChanged += new System.EventHandler(this.txtRecQuery_TextChanged);
            // 
            // picSubject
            // 
            this.picSubject.Image = ((System.Drawing.Image)(resources.GetObject("picSubject.Image")));
            this.picSubject.Location = new System.Drawing.Point(414, 101);
            this.picSubject.Name = "picSubject";
            this.picSubject.Size = new System.Drawing.Size(26, 20);
            this.picSubject.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.picSubject.TabIndex = 46;
            this.picSubject.TabStop = false;
            this.picSubject.Click += new System.EventHandler(this.picSubject_Click);
            // 
            // lblDescription
            // 
            this.lblDescription.AutoSize = true;
            this.lblDescription.Location = new System.Drawing.Point(12, 16);
            this.lblDescription.Name = "lblDescription";
            this.lblDescription.Size = new System.Drawing.Size(63, 13);
            this.lblDescription.TabIndex = 47;
            this.lblDescription.Text = "Description:";
            // 
            // txtDescription
            // 
            this.txtDescription.Location = new System.Drawing.Point(80, 13);
            this.txtDescription.Name = "txtDescription";
            this.txtDescription.Size = new System.Drawing.Size(337, 20);
            this.txtDescription.TabIndex = 48;
            this.txtDescription.TextChanged += new System.EventHandler(this.txtDescription_TextChanged);
            // 
            // flexParams
            // 
            this.flexParams.AutoClipboard = true;
            this.flexParams.ColumnInfo = resources.GetString("flexParams.ColumnInfo");
            this.flexParams.Location = new System.Drawing.Point(536, 63);
            this.flexParams.Name = "flexParams";
            this.flexParams.Rows.Count = 11;
            this.flexParams.Rows.DefaultSize = 17;
            this.flexParams.Size = new System.Drawing.Size(341, 198);
            this.flexParams.TabIndex = 49;
            this.flexParams.AfterEdit += new C1.Win.C1FlexGrid.RowColEventHandler(this.flexParams_AfterEdit);
            this.flexParams.CellButtonClick += new C1.Win.C1FlexGrid.RowColEventHandler(this.flexParams_CellButtonClick);
            this.flexParams.CellChanged += new C1.Win.C1FlexGrid.RowColEventHandler(this.flexParams_CellChanged);
            // 
            // webSubject
            // 
            this.webSubject.Location = new System.Drawing.Point(80, 101);
            this.webSubject.MinimumSize = new System.Drawing.Size(20, 20);
            this.webSubject.Name = "webSubject";
            this.webSubject.Size = new System.Drawing.Size(337, 130);
            this.webSubject.TabIndex = 45;
            // 
            // webBrowser
            // 
            this.webBrowser.Location = new System.Drawing.Point(80, 241);
            this.webBrowser.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser.Name = "webBrowser";
            this.webBrowser.Size = new System.Drawing.Size(337, 130);
            this.webBrowser.TabIndex = 41;
            // 
            // cmdTestMe
            // 
            this.cmdTestMe.Location = new System.Drawing.Point(789, 309);
            this.cmdTestMe.Name = "cmdTestMe";
            this.cmdTestMe.Size = new System.Drawing.Size(75, 62);
            this.cmdTestMe.TabIndex = 50;
            this.cmdTestMe.Text = "Test Me";
            this.cmdTestMe.UseVisualStyleBackColor = true;
            this.cmdTestMe.Click += new System.EventHandler(this.cmdTestMe_Click);
            // 
            // AutoReportDetails
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(905, 513);
            this.Controls.Add(this.cmdTestMe);
            this.Controls.Add(this.flexParams);
            this.Controls.Add(this.txtDescription);
            this.Controls.Add(this.lblDescription);
            this.Controls.Add(this.picSubject);
            this.Controls.Add(this.webSubject);
            this.Controls.Add(this.txtRecQuery);
            this.Controls.Add(this.lblRecipientQuery);
            this.Controls.Add(this.cmdQuery);
            this.Controls.Add(this.webBrowser);
            this.Controls.Add(this.txtSubject);
            this.Controls.Add(this.picMessage);
            this.Controls.Add(this.cmdCancel);
            this.Controls.Add(this.cmdSave);
            this.Controls.Add(this.lblParameters);
            this.Controls.Add(this.grRun);
            this.Controls.Add(this.grFrequency);
            this.Controls.Add(this.picReport);
            this.Controls.Add(this.picAttachment);
            this.Controls.Add(this.txtReport);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtAttachment);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtMessage);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtSender);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtRecipient);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AutoReportDetails";
            this.Text = "Report Details";
            this.Load += new System.EventHandler(this.AutoReportDetails_Load);
            ((System.ComponentModel.ISupportInitialize)(this.picAttachment)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picReport)).EndInit();
            this.grFrequency.ResumeLayout(false);
            this.grFrequency.PerformLayout();
            this.grRun.ResumeLayout(false);
            this.grRun.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picMessage)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picSubject)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.flexParams)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtRecipient;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtSender;
        private System.Windows.Forms.TextBox txtSubject;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtMessage;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtAttachment;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtReport;
        private System.Windows.Forms.PictureBox picAttachment;
        private System.Windows.Forms.PictureBox picReport;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.GroupBox grFrequency;
        private System.Windows.Forms.ComboBox cmbFrequencyType;
        private System.Windows.Forms.GroupBox grRun;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.DateTimePicker runTime;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.DateTimePicker nextRun;
        private System.Windows.Forms.Label lblParameters;
        private System.Windows.Forms.Button cmdSave;
        private System.Windows.Forms.Button cmdCancel;
        private System.Windows.Forms.ComboBox cmbFrequencyValue;
        private System.Windows.Forms.OpenFileDialog openAttachment;
        private System.Windows.Forms.OpenFileDialog openReport;
        private System.Windows.Forms.PictureBox picMessage;
        private System.Windows.Forms.WebBrowser webBrowser;
        private System.Windows.Forms.Button cmdQuery;
        private System.Windows.Forms.Label lblRecipientQuery;
        private System.Windows.Forms.TextBox txtRecQuery;
        private System.Windows.Forms.WebBrowser webSubject;
        private System.Windows.Forms.PictureBox picSubject;
        private System.Windows.Forms.Label lblDescription;
        private System.Windows.Forms.TextBox txtDescription;
        private C1.Win.C1FlexGrid.C1FlexGrid flexParams;
        private System.Windows.Forms.Button cmdTestMe;
    }
}