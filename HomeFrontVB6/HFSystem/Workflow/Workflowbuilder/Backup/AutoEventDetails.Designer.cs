namespace AutoNotice
{
    partial class AutoEventDetails
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AutoEventDetails));
            this.lblTrigger = new System.Windows.Forms.Label();
            this.txtTrigger = new System.Windows.Forms.TextBox();
            this.txtRecipient = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtMessage = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtSubject = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.grRecipients = new System.Windows.Forms.GroupBox();
            this.grOthers = new System.Windows.Forms.GroupBox();
            this.chkSellingRealtor = new System.Windows.Forms.CheckBox();
            this.chkTitleCompany = new System.Windows.Forms.CheckBox();
            this.chkBroker = new System.Windows.Forms.CheckBox();
            this.chkLender = new System.Windows.Forms.CheckBox();
            this.chkRealtor = new System.Windows.Forms.CheckBox();
            this.grSales = new System.Windows.Forms.GroupBox();
            this.chkPM = new System.Windows.Forms.CheckBox();
            this.chkBuyer = new System.Windows.Forms.CheckBox();
            this.chkSalesPerson = new System.Windows.Forms.CheckBox();
            this.chkSalesManager = new System.Windows.Forms.CheckBox();
            this.chkPurchaser = new System.Windows.Forms.CheckBox();
            this.chkDCSalesPerson = new System.Windows.Forms.CheckBox();
            this.chkEstimator = new System.Windows.Forms.CheckBox();
            this.chkCoBuyer = new System.Windows.Forms.CheckBox();
            this.grAdministrators = new System.Windows.Forms.GroupBox();
            this.chkSystemAdmin = new System.Windows.Forms.CheckBox();
            this.chkAccountingAdmin = new System.Windows.Forms.CheckBox();
            this.grVendors = new System.Windows.Forms.GroupBox();
            this.chkVendorService = new System.Windows.Forms.CheckBox();
            this.chkVendorPurchasing = new System.Windows.Forms.CheckBox();
            this.chkVendorScheduling = new System.Windows.Forms.CheckBox();
            this.picReport = new System.Windows.Forms.PictureBox();
            this.picAttachment = new System.Windows.Forms.PictureBox();
            this.txtReport = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtAttachment = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.lblParameters = new System.Windows.Forms.Label();
            this.flexParams = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.cmdSave = new System.Windows.Forms.Button();
            this.cmdCancel = new System.Windows.Forms.Button();
            this.picSubject = new System.Windows.Forms.PictureBox();
            this.picMessage = new System.Windows.Forms.PictureBox();
            this.openAttachment = new System.Windows.Forms.OpenFileDialog();
            this.openReport = new System.Windows.Forms.OpenFileDialog();
            this.chkDisabled = new System.Windows.Forms.CheckBox();
            this.webBrowser = new System.Windows.Forms.WebBrowser();
            this.lblSender = new System.Windows.Forms.Label();
            this.txtSender = new System.Windows.Forms.TextBox();
            this.grRecipients.SuspendLayout();
            this.grOthers.SuspendLayout();
            this.grSales.SuspendLayout();
            this.grAdministrators.SuspendLayout();
            this.grVendors.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picReport)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picAttachment)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.flexParams)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picSubject)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picMessage)).BeginInit();
            this.SuspendLayout();
            // 
            // lblTrigger
            // 
            this.lblTrigger.AutoSize = true;
            this.lblTrigger.Location = new System.Drawing.Point(24, 22);
            this.lblTrigger.Name = "lblTrigger";
            this.lblTrigger.Size = new System.Drawing.Size(43, 13);
            this.lblTrigger.TabIndex = 0;
            this.lblTrigger.Text = "Trigger:";
            // 
            // txtTrigger
            // 
            this.txtTrigger.Location = new System.Drawing.Point(80, 19);
            this.txtTrigger.Name = "txtTrigger";
            this.txtTrigger.ReadOnly = true;
            this.txtTrigger.Size = new System.Drawing.Size(337, 20);
            this.txtTrigger.TabIndex = 1;
            this.txtTrigger.TextChanged += new System.EventHandler(this.txtTrigger_TextChanged);
            // 
            // txtRecipient
            // 
            this.txtRecipient.Location = new System.Drawing.Point(80, 84);
            this.txtRecipient.Name = "txtRecipient";
            this.txtRecipient.Size = new System.Drawing.Size(337, 20);
            this.txtRecipient.TabIndex = 3;
            this.txtRecipient.TextChanged += new System.EventHandler(this.txtRecipient_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 87);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Recipients:";
            // 
            // txtMessage
            // 
            this.txtMessage.Location = new System.Drawing.Point(467, 379);
            this.txtMessage.Multiline = true;
            this.txtMessage.Name = "txtMessage";
            this.txtMessage.ReadOnly = true;
            this.txtMessage.Size = new System.Drawing.Size(248, 32);
            this.txtMessage.TabIndex = 11;
            this.txtMessage.Visible = false;
            this.txtMessage.TextChanged += new System.EventHandler(this.txtMessage_TextChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(9, 143);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "Message:";
            // 
            // txtSubject
            // 
            this.txtSubject.Location = new System.Drawing.Point(80, 110);
            this.txtSubject.Name = "txtSubject";
            this.txtSubject.Size = new System.Drawing.Size(337, 20);
            this.txtSubject.TabIndex = 9;
            this.txtSubject.TextChanged += new System.EventHandler(this.txtSubject_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(16, 113);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(46, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Subject:";
            // 
            // grRecipients
            // 
            this.grRecipients.Controls.Add(this.grOthers);
            this.grRecipients.Controls.Add(this.grSales);
            this.grRecipients.Controls.Add(this.grAdministrators);
            this.grRecipients.Controls.Add(this.grVendors);
            this.grRecipients.Location = new System.Drawing.Point(467, 45);
            this.grRecipients.Name = "grRecipients";
            this.grRecipients.Size = new System.Drawing.Size(248, 328);
            this.grRecipients.TabIndex = 12;
            this.grRecipients.TabStop = false;
            this.grRecipients.Text = "Recipients";
            // 
            // grOthers
            // 
            this.grOthers.Controls.Add(this.chkSellingRealtor);
            this.grOthers.Controls.Add(this.chkTitleCompany);
            this.grOthers.Controls.Add(this.chkBroker);
            this.grOthers.Controls.Add(this.chkLender);
            this.grOthers.Controls.Add(this.chkRealtor);
            this.grOthers.Location = new System.Drawing.Point(131, 98);
            this.grOthers.Name = "grOthers";
            this.grOthers.Size = new System.Drawing.Size(105, 136);
            this.grOthers.TabIndex = 19;
            this.grOthers.TabStop = false;
            this.grOthers.Text = "Others";
            // 
            // chkSellingRealtor
            // 
            this.chkSellingRealtor.AutoSize = true;
            this.chkSellingRealtor.Location = new System.Drawing.Point(6, 111);
            this.chkSellingRealtor.Name = "chkSellingRealtor";
            this.chkSellingRealtor.Size = new System.Drawing.Size(94, 17);
            this.chkSellingRealtor.TabIndex = 24;
            this.chkSellingRealtor.Text = "Selling Realtor";
            this.chkSellingRealtor.UseVisualStyleBackColor = true;
            this.chkSellingRealtor.CheckedChanged += new System.EventHandler(this.chkSellingRealtor_CheckedChanged);
            // 
            // chkTitleCompany
            // 
            this.chkTitleCompany.AutoSize = true;
            this.chkTitleCompany.Location = new System.Drawing.Point(6, 19);
            this.chkTitleCompany.Name = "chkTitleCompany";
            this.chkTitleCompany.Size = new System.Drawing.Size(93, 17);
            this.chkTitleCompany.TabIndex = 20;
            this.chkTitleCompany.Text = "Title Company";
            this.chkTitleCompany.UseVisualStyleBackColor = true;
            this.chkTitleCompany.CheckedChanged += new System.EventHandler(this.chkTitleCompany_CheckedChanged);
            // 
            // chkBroker
            // 
            this.chkBroker.AutoSize = true;
            this.chkBroker.Location = new System.Drawing.Point(6, 88);
            this.chkBroker.Name = "chkBroker";
            this.chkBroker.Size = new System.Drawing.Size(57, 17);
            this.chkBroker.TabIndex = 23;
            this.chkBroker.Text = "Broker";
            this.chkBroker.UseVisualStyleBackColor = true;
            this.chkBroker.CheckedChanged += new System.EventHandler(this.chkBroker_CheckedChanged);
            // 
            // chkLender
            // 
            this.chkLender.AutoSize = true;
            this.chkLender.Location = new System.Drawing.Point(6, 42);
            this.chkLender.Name = "chkLender";
            this.chkLender.Size = new System.Drawing.Size(59, 17);
            this.chkLender.TabIndex = 21;
            this.chkLender.Text = "Lender";
            this.chkLender.UseVisualStyleBackColor = true;
            this.chkLender.CheckedChanged += new System.EventHandler(this.chkLender_CheckedChanged);
            // 
            // chkRealtor
            // 
            this.chkRealtor.AutoSize = true;
            this.chkRealtor.Location = new System.Drawing.Point(6, 65);
            this.chkRealtor.Name = "chkRealtor";
            this.chkRealtor.Size = new System.Drawing.Size(60, 17);
            this.chkRealtor.TabIndex = 22;
            this.chkRealtor.Text = "Realtor";
            this.chkRealtor.UseVisualStyleBackColor = true;
            this.chkRealtor.CheckedChanged += new System.EventHandler(this.chkRealtor_CheckedChanged);
            // 
            // grSales
            // 
            this.grSales.Controls.Add(this.chkPM);
            this.grSales.Controls.Add(this.chkBuyer);
            this.grSales.Controls.Add(this.chkSalesPerson);
            this.grSales.Controls.Add(this.chkSalesManager);
            this.grSales.Controls.Add(this.chkPurchaser);
            this.grSales.Controls.Add(this.chkDCSalesPerson);
            this.grSales.Controls.Add(this.chkEstimator);
            this.grSales.Controls.Add(this.chkCoBuyer);
            this.grSales.Location = new System.Drawing.Point(6, 19);
            this.grSales.Name = "grSales";
            this.grSales.Size = new System.Drawing.Size(119, 201);
            this.grSales.TabIndex = 20;
            this.grSales.TabStop = false;
            this.grSales.Text = "Sales";
            // 
            // chkPM
            // 
            this.chkPM.AutoSize = true;
            this.chkPM.Location = new System.Drawing.Point(6, 131);
            this.chkPM.Name = "chkPM";
            this.chkPM.Size = new System.Drawing.Size(42, 17);
            this.chkPM.TabIndex = 3;
            this.chkPM.Text = "PM";
            this.chkPM.UseVisualStyleBackColor = true;
            this.chkPM.CheckedChanged += new System.EventHandler(this.chkPM_CheckedChanged);
            // 
            // chkBuyer
            // 
            this.chkBuyer.AutoSize = true;
            this.chkBuyer.Location = new System.Drawing.Point(6, 14);
            this.chkBuyer.Name = "chkBuyer";
            this.chkBuyer.Size = new System.Drawing.Size(53, 17);
            this.chkBuyer.TabIndex = 0;
            this.chkBuyer.Text = "Buyer";
            this.chkBuyer.UseVisualStyleBackColor = true;
            this.chkBuyer.CheckedChanged += new System.EventHandler(this.chkBuyer_CheckedChanged);
            // 
            // chkSalesPerson
            // 
            this.chkSalesPerson.AutoSize = true;
            this.chkSalesPerson.Location = new System.Drawing.Point(6, 62);
            this.chkSalesPerson.Name = "chkSalesPerson";
            this.chkSalesPerson.Size = new System.Drawing.Size(88, 17);
            this.chkSalesPerson.TabIndex = 2;
            this.chkSalesPerson.Text = "Sales Person";
            this.chkSalesPerson.UseVisualStyleBackColor = true;
            this.chkSalesPerson.CheckedChanged += new System.EventHandler(this.chkSalesPerson_CheckedChanged);
            // 
            // chkSalesManager
            // 
            this.chkSalesManager.AutoSize = true;
            this.chkSalesManager.Location = new System.Drawing.Point(6, 108);
            this.chkSalesManager.Name = "chkSalesManager";
            this.chkSalesManager.Size = new System.Drawing.Size(97, 17);
            this.chkSalesManager.TabIndex = 13;
            this.chkSalesManager.Text = "Sales Manager";
            this.chkSalesManager.UseVisualStyleBackColor = true;
            this.chkSalesManager.CheckedChanged += new System.EventHandler(this.chkSalesManager_CheckedChanged);
            // 
            // chkPurchaser
            // 
            this.chkPurchaser.AutoSize = true;
            this.chkPurchaser.Location = new System.Drawing.Point(6, 177);
            this.chkPurchaser.Name = "chkPurchaser";
            this.chkPurchaser.Size = new System.Drawing.Size(74, 17);
            this.chkPurchaser.TabIndex = 15;
            this.chkPurchaser.Text = "Purchaser";
            this.chkPurchaser.UseVisualStyleBackColor = true;
            this.chkPurchaser.CheckedChanged += new System.EventHandler(this.chkPurchaser_CheckedChanged);
            // 
            // chkDCSalesPerson
            // 
            this.chkDCSalesPerson.AutoSize = true;
            this.chkDCSalesPerson.Location = new System.Drawing.Point(6, 85);
            this.chkDCSalesPerson.Name = "chkDCSalesPerson";
            this.chkDCSalesPerson.Size = new System.Drawing.Size(106, 17);
            this.chkDCSalesPerson.TabIndex = 16;
            this.chkDCSalesPerson.Text = "DC Sales Person";
            this.chkDCSalesPerson.UseVisualStyleBackColor = true;
            this.chkDCSalesPerson.CheckedChanged += new System.EventHandler(this.chkDCSalesPerson_CheckedChanged);
            // 
            // chkEstimator
            // 
            this.chkEstimator.AutoSize = true;
            this.chkEstimator.Location = new System.Drawing.Point(6, 154);
            this.chkEstimator.Name = "chkEstimator";
            this.chkEstimator.Size = new System.Drawing.Size(69, 17);
            this.chkEstimator.TabIndex = 14;
            this.chkEstimator.Text = "Estimator";
            this.chkEstimator.UseVisualStyleBackColor = true;
            this.chkEstimator.CheckedChanged += new System.EventHandler(this.chkEstimator_CheckedChanged);
            // 
            // chkCoBuyer
            // 
            this.chkCoBuyer.AutoSize = true;
            this.chkCoBuyer.Location = new System.Drawing.Point(6, 38);
            this.chkCoBuyer.Name = "chkCoBuyer";
            this.chkCoBuyer.Size = new System.Drawing.Size(69, 17);
            this.chkCoBuyer.TabIndex = 1;
            this.chkCoBuyer.Text = "Co-Buyer";
            this.chkCoBuyer.UseVisualStyleBackColor = true;
            this.chkCoBuyer.CheckedChanged += new System.EventHandler(this.chkCoBuyer_CheckedChanged);
            // 
            // grAdministrators
            // 
            this.grAdministrators.Controls.Add(this.chkSystemAdmin);
            this.grAdministrators.Controls.Add(this.chkAccountingAdmin);
            this.grAdministrators.Location = new System.Drawing.Point(131, 19);
            this.grAdministrators.Name = "grAdministrators";
            this.grAdministrators.Size = new System.Drawing.Size(105, 73);
            this.grAdministrators.TabIndex = 18;
            this.grAdministrators.TabStop = false;
            this.grAdministrators.Text = "Administrators";
            // 
            // chkSystemAdmin
            // 
            this.chkSystemAdmin.AutoSize = true;
            this.chkSystemAdmin.Location = new System.Drawing.Point(6, 46);
            this.chkSystemAdmin.Name = "chkSystemAdmin";
            this.chkSystemAdmin.Size = new System.Drawing.Size(60, 17);
            this.chkSystemAdmin.TabIndex = 20;
            this.chkSystemAdmin.Text = "System";
            this.chkSystemAdmin.UseVisualStyleBackColor = true;
            this.chkSystemAdmin.CheckedChanged += new System.EventHandler(this.chkSystemAdmin_CheckedChanged);
            // 
            // chkAccountingAdmin
            // 
            this.chkAccountingAdmin.AutoSize = true;
            this.chkAccountingAdmin.Location = new System.Drawing.Point(6, 23);
            this.chkAccountingAdmin.Name = "chkAccountingAdmin";
            this.chkAccountingAdmin.Size = new System.Drawing.Size(80, 17);
            this.chkAccountingAdmin.TabIndex = 19;
            this.chkAccountingAdmin.Text = "Accounting";
            this.chkAccountingAdmin.UseVisualStyleBackColor = true;
            this.chkAccountingAdmin.CheckedChanged += new System.EventHandler(this.chkAccountingAdmin_CheckedChanged);
            // 
            // grVendors
            // 
            this.grVendors.Controls.Add(this.chkVendorService);
            this.grVendors.Controls.Add(this.chkVendorPurchasing);
            this.grVendors.Controls.Add(this.chkVendorScheduling);
            this.grVendors.Location = new System.Drawing.Point(6, 226);
            this.grVendors.Name = "grVendors";
            this.grVendors.Size = new System.Drawing.Size(95, 92);
            this.grVendors.TabIndex = 17;
            this.grVendors.TabStop = false;
            this.grVendors.Text = "Vendors";
            // 
            // chkVendorService
            // 
            this.chkVendorService.AutoSize = true;
            this.chkVendorService.Location = new System.Drawing.Point(6, 65);
            this.chkVendorService.Name = "chkVendorService";
            this.chkVendorService.Size = new System.Drawing.Size(62, 17);
            this.chkVendorService.TabIndex = 20;
            this.chkVendorService.Text = "Service";
            this.chkVendorService.UseVisualStyleBackColor = true;
            this.chkVendorService.CheckedChanged += new System.EventHandler(this.chkVendorService_CheckedChanged);
            // 
            // chkVendorPurchasing
            // 
            this.chkVendorPurchasing.AutoSize = true;
            this.chkVendorPurchasing.Location = new System.Drawing.Point(6, 19);
            this.chkVendorPurchasing.Name = "chkVendorPurchasing";
            this.chkVendorPurchasing.Size = new System.Drawing.Size(79, 17);
            this.chkVendorPurchasing.TabIndex = 18;
            this.chkVendorPurchasing.Text = "Purchasing";
            this.chkVendorPurchasing.UseVisualStyleBackColor = true;
            this.chkVendorPurchasing.CheckedChanged += new System.EventHandler(this.chkVendorPurchasing_CheckedChanged);
            // 
            // chkVendorScheduling
            // 
            this.chkVendorScheduling.AutoSize = true;
            this.chkVendorScheduling.Location = new System.Drawing.Point(6, 42);
            this.chkVendorScheduling.Name = "chkVendorScheduling";
            this.chkVendorScheduling.Size = new System.Drawing.Size(79, 17);
            this.chkVendorScheduling.TabIndex = 19;
            this.chkVendorScheduling.Text = "Scheduling";
            this.chkVendorScheduling.UseVisualStyleBackColor = true;
            this.chkVendorScheduling.CheckedChanged += new System.EventHandler(this.chkVendorScheduling_CheckedChanged);
            // 
            // picReport
            // 
            this.picReport.Image = ((System.Drawing.Image)(resources.GetObject("picReport.Image")));
            this.picReport.Location = new System.Drawing.Point(414, 318);
            this.picReport.Name = "picReport";
            this.picReport.Size = new System.Drawing.Size(26, 20);
            this.picReport.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.picReport.TabIndex = 31;
            this.picReport.TabStop = false;
            this.picReport.Click += new System.EventHandler(this.picReport_Click);
            // 
            // picAttachment
            // 
            this.picAttachment.Image = ((System.Drawing.Image)(resources.GetObject("picAttachment.Image")));
            this.picAttachment.Location = new System.Drawing.Point(414, 290);
            this.picAttachment.Name = "picAttachment";
            this.picAttachment.Size = new System.Drawing.Size(26, 20);
            this.picAttachment.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.picAttachment.TabIndex = 30;
            this.picAttachment.TabStop = false;
            this.picAttachment.Click += new System.EventHandler(this.picAttachment_Click);
            // 
            // txtReport
            // 
            this.txtReport.Location = new System.Drawing.Point(80, 318);
            this.txtReport.Name = "txtReport";
            this.txtReport.Size = new System.Drawing.Size(337, 20);
            this.txtReport.TabIndex = 29;
            this.txtReport.TextChanged += new System.EventHandler(this.txtReport_TextChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(23, 318);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(39, 13);
            this.label6.TabIndex = 28;
            this.label6.Text = "Report";
            // 
            // txtAttachment
            // 
            this.txtAttachment.Location = new System.Drawing.Point(80, 290);
            this.txtAttachment.Name = "txtAttachment";
            this.txtAttachment.Size = new System.Drawing.Size(337, 20);
            this.txtAttachment.TabIndex = 27;
            this.txtAttachment.TextChanged += new System.EventHandler(this.txtAttachment_TextChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(-2, 290);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(64, 13);
            this.label5.TabIndex = 26;
            this.label5.Text = "Attachment:";
            // 
            // lblParameters
            // 
            this.lblParameters.AutoSize = true;
            this.lblParameters.Location = new System.Drawing.Point(7, 356);
            this.lblParameters.Name = "lblParameters";
            this.lblParameters.Size = new System.Drawing.Size(60, 13);
            this.lblParameters.TabIndex = 33;
            this.lblParameters.Text = "Parameters";
            // 
            // flexParams
            // 
            this.flexParams.AutoClipboard = true;
            this.flexParams.ColumnInfo = resources.GetString("flexParams.ColumnInfo");
            this.flexParams.Location = new System.Drawing.Point(80, 356);
            this.flexParams.Name = "flexParams";
            this.flexParams.Rows.Count = 11;
            this.flexParams.Rows.DefaultSize = 17;
            this.flexParams.Size = new System.Drawing.Size(337, 205);
            this.flexParams.TabIndex = 34;
            this.flexParams.AfterEdit += new C1.Win.C1FlexGrid.RowColEventHandler(this.flexParams_AfterEdit);
            this.flexParams.CellButtonClick += new C1.Win.C1FlexGrid.RowColEventHandler(this.flexParams_CellButtonClick);
            this.flexParams.DoubleClick += new System.EventHandler(this.flexParams_DoubleClick);
            this.flexParams.CellChanged += new C1.Win.C1FlexGrid.RowColEventHandler(this.flexParams_CellChanged);
            // 
            // cmdSave
            // 
            this.cmdSave.Location = new System.Drawing.Point(559, 538);
            this.cmdSave.Name = "cmdSave";
            this.cmdSave.Size = new System.Drawing.Size(75, 23);
            this.cmdSave.TabIndex = 35;
            this.cmdSave.Text = "Save";
            this.cmdSave.UseVisualStyleBackColor = true;
            this.cmdSave.Click += new System.EventHandler(this.cmdSave_Click);
            // 
            // cmdCancel
            // 
            this.cmdCancel.Location = new System.Drawing.Point(640, 538);
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.Size = new System.Drawing.Size(75, 23);
            this.cmdCancel.TabIndex = 36;
            this.cmdCancel.Text = "Cancel";
            this.cmdCancel.UseVisualStyleBackColor = true;
            this.cmdCancel.Click += new System.EventHandler(this.cmdCancel_Click);
            // 
            // picSubject
            // 
            this.picSubject.Image = ((System.Drawing.Image)(resources.GetObject("picSubject.Image")));
            this.picSubject.Location = new System.Drawing.Point(414, 110);
            this.picSubject.Name = "picSubject";
            this.picSubject.Size = new System.Drawing.Size(26, 20);
            this.picSubject.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.picSubject.TabIndex = 37;
            this.picSubject.TabStop = false;
            this.picSubject.Click += new System.EventHandler(this.picSubject_Click);
            // 
            // picMessage
            // 
            this.picMessage.Image = ((System.Drawing.Image)(resources.GetObject("picMessage.Image")));
            this.picMessage.Location = new System.Drawing.Point(414, 140);
            this.picMessage.Name = "picMessage";
            this.picMessage.Size = new System.Drawing.Size(26, 20);
            this.picMessage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.picMessage.TabIndex = 38;
            this.picMessage.TabStop = false;
            this.picMessage.Click += new System.EventHandler(this.picMessage_Click);
            // 
            // openReport
            // 
            this.openReport.Filter = "(*.rpt)|*.rpt";
            // 
            // chkDisabled
            // 
            this.chkDisabled.AutoSize = true;
            this.chkDisabled.Location = new System.Drawing.Point(479, 21);
            this.chkDisabled.Name = "chkDisabled";
            this.chkDisabled.Size = new System.Drawing.Size(67, 17);
            this.chkDisabled.TabIndex = 39;
            this.chkDisabled.Text = "Disabled";
            this.chkDisabled.UseVisualStyleBackColor = true;
            this.chkDisabled.CheckedChanged += new System.EventHandler(this.chkDisabled_CheckedChanged);
            // 
            // webBrowser
            // 
            this.webBrowser.Location = new System.Drawing.Point(80, 140);
            this.webBrowser.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser.Name = "webBrowser";
            this.webBrowser.Size = new System.Drawing.Size(337, 134);
            this.webBrowser.TabIndex = 40;
            // 
            // lblSender
            // 
            this.lblSender.AutoSize = true;
            this.lblSender.Location = new System.Drawing.Point(23, 54);
            this.lblSender.Name = "lblSender";
            this.lblSender.Size = new System.Drawing.Size(44, 13);
            this.lblSender.TabIndex = 41;
            this.lblSender.Text = "Sender:";
            // 
            // txtSender
            // 
            this.txtSender.Location = new System.Drawing.Point(80, 51);
            this.txtSender.Name = "txtSender";
            this.txtSender.Size = new System.Drawing.Size(337, 20);
            this.txtSender.TabIndex = 42;
            this.txtSender.TextChanged += new System.EventHandler(this.txtSender_TextChanged);
            // 
            // AutoEventDetails
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(731, 594);
            this.Controls.Add(this.txtSender);
            this.Controls.Add(this.lblSender);
            this.Controls.Add(this.webBrowser);
            this.Controls.Add(this.chkDisabled);
            this.Controls.Add(this.picMessage);
            this.Controls.Add(this.picSubject);
            this.Controls.Add(this.cmdCancel);
            this.Controls.Add(this.cmdSave);
            this.Controls.Add(this.flexParams);
            this.Controls.Add(this.lblParameters);
            this.Controls.Add(this.picReport);
            this.Controls.Add(this.picAttachment);
            this.Controls.Add(this.txtReport);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtAttachment);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.grRecipients);
            this.Controls.Add(this.txtMessage);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtSubject);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtRecipient);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtTrigger);
            this.Controls.Add(this.lblTrigger);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AutoEventDetails";
            this.Text = "Trigger Details";
            this.Load += new System.EventHandler(this.AutoEventDetails_Load);
            this.grRecipients.ResumeLayout(false);
            this.grOthers.ResumeLayout(false);
            this.grOthers.PerformLayout();
            this.grSales.ResumeLayout(false);
            this.grSales.PerformLayout();
            this.grAdministrators.ResumeLayout(false);
            this.grAdministrators.PerformLayout();
            this.grVendors.ResumeLayout(false);
            this.grVendors.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picReport)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picAttachment)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.flexParams)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picSubject)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picMessage)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblTrigger;
        private System.Windows.Forms.TextBox txtTrigger;
        private System.Windows.Forms.TextBox txtRecipient;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtMessage;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtSubject;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.GroupBox grRecipients;
        private System.Windows.Forms.CheckBox chkCoBuyer;
        private System.Windows.Forms.CheckBox chkBuyer;
        private System.Windows.Forms.CheckBox chkSalesPerson;
        private System.Windows.Forms.CheckBox chkPM;
        private System.Windows.Forms.CheckBox chkSalesManager;
        private System.Windows.Forms.CheckBox chkEstimator;
        private System.Windows.Forms.CheckBox chkPurchaser;
        private System.Windows.Forms.CheckBox chkDCSalesPerson;
        private System.Windows.Forms.GroupBox grVendors;
        private System.Windows.Forms.CheckBox chkVendorService;
        private System.Windows.Forms.CheckBox chkVendorPurchasing;
        private System.Windows.Forms.CheckBox chkVendorScheduling;
        private System.Windows.Forms.GroupBox grAdministrators;
        private System.Windows.Forms.CheckBox chkSystemAdmin;
        private System.Windows.Forms.CheckBox chkAccountingAdmin;
        private System.Windows.Forms.GroupBox grOthers;
        private System.Windows.Forms.CheckBox chkTitleCompany;
        private System.Windows.Forms.CheckBox chkLender;
        private System.Windows.Forms.CheckBox chkRealtor;
        private System.Windows.Forms.CheckBox chkBroker;
        private System.Windows.Forms.CheckBox chkSellingRealtor;
        private System.Windows.Forms.GroupBox grSales;
        private System.Windows.Forms.PictureBox picReport;
        private System.Windows.Forms.PictureBox picAttachment;
        private System.Windows.Forms.TextBox txtReport;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtAttachment;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label lblParameters;
        private C1.Win.C1FlexGrid.C1FlexGrid flexParams;
        private System.Windows.Forms.Button cmdSave;
        private System.Windows.Forms.Button cmdCancel;
        private System.Windows.Forms.PictureBox picSubject;
        private System.Windows.Forms.PictureBox picMessage;
        private System.Windows.Forms.OpenFileDialog openAttachment;
        private System.Windows.Forms.OpenFileDialog openReport;
        private System.Windows.Forms.CheckBox chkDisabled;
        private System.Windows.Forms.WebBrowser webBrowser;
        private System.Windows.Forms.Label lblSender;
        private System.Windows.Forms.TextBox txtSender;
    }
}