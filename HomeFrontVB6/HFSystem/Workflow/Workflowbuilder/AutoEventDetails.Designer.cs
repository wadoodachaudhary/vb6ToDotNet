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
            this.txtCC = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
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
            this.lblTrigger.Location = new System.Drawing.Point(32, 27);
            this.lblTrigger.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblTrigger.Name = "lblTrigger";
            this.lblTrigger.Size = new System.Drawing.Size(58, 17);
            this.lblTrigger.TabIndex = 0;
            this.lblTrigger.Text = "Trigger:";
            this.lblTrigger.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // txtTrigger
            // 
            this.txtTrigger.Location = new System.Drawing.Point(107, 23);
            this.txtTrigger.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtTrigger.Name = "txtTrigger";
            this.txtTrigger.ReadOnly = true;
            this.txtTrigger.Size = new System.Drawing.Size(448, 22);
            this.txtTrigger.TabIndex = 0;
            this.txtTrigger.TextChanged += new System.EventHandler(this.txtTrigger_TextChanged);
            // 
            // txtRecipient
            // 
            this.txtRecipient.Location = new System.Drawing.Point(107, 78);
            this.txtRecipient.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtRecipient.Name = "txtRecipient";
            this.txtRecipient.Size = new System.Drawing.Size(448, 22);
            this.txtRecipient.TabIndex = 2;
            this.txtRecipient.TextChanged += new System.EventHandler(this.txtRecipient_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 83);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 17);
            this.label1.TabIndex = 2;
            this.label1.Text = "Recipients:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // txtMessage
            // 
            this.txtMessage.Location = new System.Drawing.Point(623, 466);
            this.txtMessage.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtMessage.Multiline = true;
            this.txtMessage.Name = "txtMessage";
            this.txtMessage.ReadOnly = true;
            this.txtMessage.Size = new System.Drawing.Size(329, 38);
            this.txtMessage.TabIndex = 11;
            this.txtMessage.Visible = false;
            this.txtMessage.TextChanged += new System.EventHandler(this.txtMessage_TextChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(21, 176);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(69, 17);
            this.label4.TabIndex = 10;
            this.label4.Text = "Message:";
            this.label4.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // txtSubject
            // 
            this.txtSubject.Location = new System.Drawing.Point(107, 135);
            this.txtSubject.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtSubject.Name = "txtSubject";
            this.txtSubject.Size = new System.Drawing.Size(448, 22);
            this.txtSubject.TabIndex = 4;
            this.txtSubject.TextChanged += new System.EventHandler(this.txtSubject_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(31, 138);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 17);
            this.label3.TabIndex = 8;
            this.label3.Text = "Subject:";
            this.label3.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // grRecipients
            // 
            this.grRecipients.Controls.Add(this.grOthers);
            this.grRecipients.Controls.Add(this.grSales);
            this.grRecipients.Controls.Add(this.grAdministrators);
            this.grRecipients.Controls.Add(this.grVendors);
            this.grRecipients.Location = new System.Drawing.Point(623, 55);
            this.grRecipients.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.grRecipients.Name = "grRecipients";
            this.grRecipients.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.grRecipients.Size = new System.Drawing.Size(331, 404);
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
            this.grOthers.Location = new System.Drawing.Point(175, 121);
            this.grOthers.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.grOthers.Name = "grOthers";
            this.grOthers.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.grOthers.Size = new System.Drawing.Size(140, 167);
            this.grOthers.TabIndex = 22;
            this.grOthers.TabStop = false;
            this.grOthers.Text = "Others";
            // 
            // chkSellingRealtor
            // 
            this.chkSellingRealtor.AutoSize = true;
            this.chkSellingRealtor.Location = new System.Drawing.Point(8, 137);
            this.chkSellingRealtor.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chkSellingRealtor.Name = "chkSellingRealtor";
            this.chkSellingRealtor.Size = new System.Drawing.Size(122, 21);
            this.chkSellingRealtor.TabIndex = 23;
            this.chkSellingRealtor.Text = "Selling Realtor";
            this.chkSellingRealtor.UseVisualStyleBackColor = true;
            this.chkSellingRealtor.CheckedChanged += new System.EventHandler(this.chkSellingRealtor_CheckedChanged);
            // 
            // chkTitleCompany
            // 
            this.chkTitleCompany.AutoSize = true;
            this.chkTitleCompany.Location = new System.Drawing.Point(8, 23);
            this.chkTitleCompany.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chkTitleCompany.Name = "chkTitleCompany";
            this.chkTitleCompany.Size = new System.Drawing.Size(120, 21);
            this.chkTitleCompany.TabIndex = 19;
            this.chkTitleCompany.Text = "Title Company";
            this.chkTitleCompany.UseVisualStyleBackColor = true;
            this.chkTitleCompany.CheckedChanged += new System.EventHandler(this.chkTitleCompany_CheckedChanged);
            // 
            // chkBroker
            // 
            this.chkBroker.AutoSize = true;
            this.chkBroker.Location = new System.Drawing.Point(8, 108);
            this.chkBroker.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chkBroker.Name = "chkBroker";
            this.chkBroker.Size = new System.Drawing.Size(72, 21);
            this.chkBroker.TabIndex = 22;
            this.chkBroker.Text = "Broker";
            this.chkBroker.UseVisualStyleBackColor = true;
            this.chkBroker.CheckedChanged += new System.EventHandler(this.chkBroker_CheckedChanged);
            // 
            // chkLender
            // 
            this.chkLender.AutoSize = true;
            this.chkLender.Location = new System.Drawing.Point(8, 52);
            this.chkLender.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chkLender.Name = "chkLender";
            this.chkLender.Size = new System.Drawing.Size(75, 21);
            this.chkLender.TabIndex = 20;
            this.chkLender.Text = "Lender";
            this.chkLender.UseVisualStyleBackColor = true;
            this.chkLender.CheckedChanged += new System.EventHandler(this.chkLender_CheckedChanged);
            // 
            // chkRealtor
            // 
            this.chkRealtor.AutoSize = true;
            this.chkRealtor.Location = new System.Drawing.Point(8, 80);
            this.chkRealtor.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chkRealtor.Name = "chkRealtor";
            this.chkRealtor.Size = new System.Drawing.Size(76, 21);
            this.chkRealtor.TabIndex = 21;
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
            this.grSales.Location = new System.Drawing.Point(8, 23);
            this.grSales.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.grSales.Name = "grSales";
            this.grSales.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.grSales.Size = new System.Drawing.Size(159, 247);
            this.grSales.TabIndex = 20;
            this.grSales.TabStop = false;
            this.grSales.Text = "Sales";
            // 
            // chkPM
            // 
            this.chkPM.AutoSize = true;
            this.chkPM.Location = new System.Drawing.Point(8, 161);
            this.chkPM.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chkPM.Name = "chkPM";
            this.chkPM.Size = new System.Drawing.Size(50, 21);
            this.chkPM.TabIndex = 14;
            this.chkPM.Text = "PM";
            this.chkPM.UseVisualStyleBackColor = true;
            this.chkPM.CheckedChanged += new System.EventHandler(this.chkPM_CheckedChanged);
            // 
            // chkBuyer
            // 
            this.chkBuyer.AutoSize = true;
            this.chkBuyer.Location = new System.Drawing.Point(8, 17);
            this.chkBuyer.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chkBuyer.Name = "chkBuyer";
            this.chkBuyer.Size = new System.Drawing.Size(67, 21);
            this.chkBuyer.TabIndex = 9;
            this.chkBuyer.Text = "Buyer";
            this.chkBuyer.UseVisualStyleBackColor = true;
            this.chkBuyer.CheckedChanged += new System.EventHandler(this.chkBuyer_CheckedChanged);
            // 
            // chkSalesPerson
            // 
            this.chkSalesPerson.AutoSize = true;
            this.chkSalesPerson.Location = new System.Drawing.Point(8, 76);
            this.chkSalesPerson.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chkSalesPerson.Name = "chkSalesPerson";
            this.chkSalesPerson.Size = new System.Drawing.Size(114, 21);
            this.chkSalesPerson.TabIndex = 11;
            this.chkSalesPerson.Text = "Sales Person";
            this.chkSalesPerson.UseVisualStyleBackColor = true;
            this.chkSalesPerson.CheckedChanged += new System.EventHandler(this.chkSalesPerson_CheckedChanged);
            // 
            // chkSalesManager
            // 
            this.chkSalesManager.AutoSize = true;
            this.chkSalesManager.Location = new System.Drawing.Point(8, 133);
            this.chkSalesManager.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chkSalesManager.Name = "chkSalesManager";
            this.chkSalesManager.Size = new System.Drawing.Size(125, 21);
            this.chkSalesManager.TabIndex = 13;
            this.chkSalesManager.Text = "Sales Manager";
            this.chkSalesManager.UseVisualStyleBackColor = true;
            this.chkSalesManager.CheckedChanged += new System.EventHandler(this.chkSalesManager_CheckedChanged);
            // 
            // chkPurchaser
            // 
            this.chkPurchaser.AutoSize = true;
            this.chkPurchaser.Location = new System.Drawing.Point(8, 218);
            this.chkPurchaser.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chkPurchaser.Name = "chkPurchaser";
            this.chkPurchaser.Size = new System.Drawing.Size(95, 21);
            this.chkPurchaser.TabIndex = 16;
            this.chkPurchaser.Text = "Purchaser";
            this.chkPurchaser.UseVisualStyleBackColor = true;
            this.chkPurchaser.CheckedChanged += new System.EventHandler(this.chkPurchaser_CheckedChanged);
            // 
            // chkDCSalesPerson
            // 
            this.chkDCSalesPerson.AutoSize = true;
            this.chkDCSalesPerson.Location = new System.Drawing.Point(8, 105);
            this.chkDCSalesPerson.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chkDCSalesPerson.Name = "chkDCSalesPerson";
            this.chkDCSalesPerson.Size = new System.Drawing.Size(137, 21);
            this.chkDCSalesPerson.TabIndex = 12;
            this.chkDCSalesPerson.Text = "DC Sales Person";
            this.chkDCSalesPerson.UseVisualStyleBackColor = true;
            this.chkDCSalesPerson.CheckedChanged += new System.EventHandler(this.chkDCSalesPerson_CheckedChanged);
            // 
            // chkEstimator
            // 
            this.chkEstimator.AutoSize = true;
            this.chkEstimator.Location = new System.Drawing.Point(8, 190);
            this.chkEstimator.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chkEstimator.Name = "chkEstimator";
            this.chkEstimator.Size = new System.Drawing.Size(89, 21);
            this.chkEstimator.TabIndex = 15;
            this.chkEstimator.Text = "Estimator";
            this.chkEstimator.UseVisualStyleBackColor = true;
            this.chkEstimator.CheckedChanged += new System.EventHandler(this.chkEstimator_CheckedChanged);
            // 
            // chkCoBuyer
            // 
            this.chkCoBuyer.AutoSize = true;
            this.chkCoBuyer.Location = new System.Drawing.Point(8, 47);
            this.chkCoBuyer.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chkCoBuyer.Name = "chkCoBuyer";
            this.chkCoBuyer.Size = new System.Drawing.Size(89, 21);
            this.chkCoBuyer.TabIndex = 10;
            this.chkCoBuyer.Text = "Co-Buyer";
            this.chkCoBuyer.UseVisualStyleBackColor = true;
            this.chkCoBuyer.CheckedChanged += new System.EventHandler(this.chkCoBuyer_CheckedChanged);
            // 
            // grAdministrators
            // 
            this.grAdministrators.Controls.Add(this.chkSystemAdmin);
            this.grAdministrators.Controls.Add(this.chkAccountingAdmin);
            this.grAdministrators.Location = new System.Drawing.Point(175, 23);
            this.grAdministrators.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.grAdministrators.Name = "grAdministrators";
            this.grAdministrators.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.grAdministrators.Size = new System.Drawing.Size(140, 90);
            this.grAdministrators.TabIndex = 18;
            this.grAdministrators.TabStop = false;
            this.grAdministrators.Text = "Administrators";
            // 
            // chkSystemAdmin
            // 
            this.chkSystemAdmin.AutoSize = true;
            this.chkSystemAdmin.Location = new System.Drawing.Point(8, 57);
            this.chkSystemAdmin.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chkSystemAdmin.Name = "chkSystemAdmin";
            this.chkSystemAdmin.Size = new System.Drawing.Size(76, 21);
            this.chkSystemAdmin.TabIndex = 18;
            this.chkSystemAdmin.Text = "System";
            this.chkSystemAdmin.UseVisualStyleBackColor = true;
            this.chkSystemAdmin.CheckedChanged += new System.EventHandler(this.chkSystemAdmin_CheckedChanged);
            // 
            // chkAccountingAdmin
            // 
            this.chkAccountingAdmin.AutoSize = true;
            this.chkAccountingAdmin.Location = new System.Drawing.Point(8, 28);
            this.chkAccountingAdmin.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chkAccountingAdmin.Name = "chkAccountingAdmin";
            this.chkAccountingAdmin.Size = new System.Drawing.Size(100, 21);
            this.chkAccountingAdmin.TabIndex = 17;
            this.chkAccountingAdmin.Text = "Accounting";
            this.chkAccountingAdmin.UseVisualStyleBackColor = true;
            this.chkAccountingAdmin.CheckedChanged += new System.EventHandler(this.chkAccountingAdmin_CheckedChanged);
            // 
            // grVendors
            // 
            this.grVendors.Controls.Add(this.chkVendorService);
            this.grVendors.Controls.Add(this.chkVendorPurchasing);
            this.grVendors.Controls.Add(this.chkVendorScheduling);
            this.grVendors.Location = new System.Drawing.Point(8, 278);
            this.grVendors.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.grVendors.Name = "grVendors";
            this.grVendors.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.grVendors.Size = new System.Drawing.Size(127, 113);
            this.grVendors.TabIndex = 17;
            this.grVendors.TabStop = false;
            this.grVendors.Text = "Vendors";
            // 
            // chkVendorService
            // 
            this.chkVendorService.AutoSize = true;
            this.chkVendorService.Location = new System.Drawing.Point(8, 80);
            this.chkVendorService.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chkVendorService.Name = "chkVendorService";
            this.chkVendorService.Size = new System.Drawing.Size(77, 21);
            this.chkVendorService.TabIndex = 26;
            this.chkVendorService.Text = "Service";
            this.chkVendorService.UseVisualStyleBackColor = true;
            this.chkVendorService.CheckedChanged += new System.EventHandler(this.chkVendorService_CheckedChanged);
            // 
            // chkVendorPurchasing
            // 
            this.chkVendorPurchasing.AutoSize = true;
            this.chkVendorPurchasing.Location = new System.Drawing.Point(8, 23);
            this.chkVendorPurchasing.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chkVendorPurchasing.Name = "chkVendorPurchasing";
            this.chkVendorPurchasing.Size = new System.Drawing.Size(101, 21);
            this.chkVendorPurchasing.TabIndex = 24;
            this.chkVendorPurchasing.Text = "Purchasing";
            this.chkVendorPurchasing.UseVisualStyleBackColor = true;
            this.chkVendorPurchasing.CheckedChanged += new System.EventHandler(this.chkVendorPurchasing_CheckedChanged);
            // 
            // chkVendorScheduling
            // 
            this.chkVendorScheduling.AutoSize = true;
            this.chkVendorScheduling.Location = new System.Drawing.Point(8, 52);
            this.chkVendorScheduling.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chkVendorScheduling.Name = "chkVendorScheduling";
            this.chkVendorScheduling.Size = new System.Drawing.Size(100, 21);
            this.chkVendorScheduling.TabIndex = 25;
            this.chkVendorScheduling.Text = "Scheduling";
            this.chkVendorScheduling.UseVisualStyleBackColor = true;
            this.chkVendorScheduling.CheckedChanged += new System.EventHandler(this.chkVendorScheduling_CheckedChanged);
            // 
            // picReport
            // 
            this.picReport.Image = ((System.Drawing.Image)(resources.GetObject("picReport.Image")));
            this.picReport.Location = new System.Drawing.Point(552, 391);
            this.picReport.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.picReport.Name = "picReport";
            this.picReport.Size = new System.Drawing.Size(35, 25);
            this.picReport.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.picReport.TabIndex = 31;
            this.picReport.TabStop = false;
            this.picReport.Click += new System.EventHandler(this.picReport_Click);
            // 
            // picAttachment
            // 
            this.picAttachment.Image = ((System.Drawing.Image)(resources.GetObject("picAttachment.Image")));
            this.picAttachment.Location = new System.Drawing.Point(552, 357);
            this.picAttachment.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.picAttachment.Name = "picAttachment";
            this.picAttachment.Size = new System.Drawing.Size(35, 25);
            this.picAttachment.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.picAttachment.TabIndex = 30;
            this.picAttachment.TabStop = false;
            this.picAttachment.Click += new System.EventHandler(this.picAttachment_Click);
            // 
            // txtReport
            // 
            this.txtReport.Location = new System.Drawing.Point(107, 391);
            this.txtReport.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtReport.Name = "txtReport";
            this.txtReport.Size = new System.Drawing.Size(448, 22);
            this.txtReport.TabIndex = 7;
            this.txtReport.TextChanged += new System.EventHandler(this.txtReport_TextChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(39, 391);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(51, 17);
            this.label6.TabIndex = 28;
            this.label6.Text = "Report";
            this.label6.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // txtAttachment
            // 
            this.txtAttachment.Location = new System.Drawing.Point(107, 357);
            this.txtAttachment.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtAttachment.Name = "txtAttachment";
            this.txtAttachment.Size = new System.Drawing.Size(448, 22);
            this.txtAttachment.TabIndex = 6;
            this.txtAttachment.TextChanged += new System.EventHandler(this.txtAttachment_TextChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(7, 357);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(83, 17);
            this.label5.TabIndex = 26;
            this.label5.Text = "Attachment:";
            this.label5.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // lblParameters
            // 
            this.lblParameters.AutoSize = true;
            this.lblParameters.Location = new System.Drawing.Point(9, 438);
            this.lblParameters.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblParameters.Name = "lblParameters";
            this.lblParameters.Size = new System.Drawing.Size(81, 17);
            this.lblParameters.TabIndex = 33;
            this.lblParameters.Text = "Parameters";
            this.lblParameters.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // flexParams
            // 
            this.flexParams.AutoClipboard = true;
            this.flexParams.ColumnInfo = resources.GetString("flexParams.ColumnInfo");
            this.flexParams.Location = new System.Drawing.Point(107, 438);
            this.flexParams.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.flexParams.Name = "flexParams";
            this.flexParams.Rows.Count = 11;
            this.flexParams.Rows.DefaultSize = 17;
            this.flexParams.Size = new System.Drawing.Size(448, 251);
            this.flexParams.TabIndex = 27;
            this.flexParams.AfterEdit += new C1.Win.C1FlexGrid.RowColEventHandler(this.flexParams_AfterEdit);
            this.flexParams.CellButtonClick += new C1.Win.C1FlexGrid.RowColEventHandler(this.flexParams_CellButtonClick);
            this.flexParams.CellChanged += new C1.Win.C1FlexGrid.RowColEventHandler(this.flexParams_CellChanged);
            this.flexParams.DoubleClick += new System.EventHandler(this.flexParams_DoubleClick);
            // 
            // cmdSave
            // 
            this.cmdSave.Location = new System.Drawing.Point(745, 662);
            this.cmdSave.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cmdSave.Name = "cmdSave";
            this.cmdSave.Size = new System.Drawing.Size(100, 28);
            this.cmdSave.TabIndex = 35;
            this.cmdSave.Text = "Save";
            this.cmdSave.UseVisualStyleBackColor = true;
            this.cmdSave.Click += new System.EventHandler(this.cmdSave_Click);
            // 
            // cmdCancel
            // 
            this.cmdCancel.Location = new System.Drawing.Point(853, 662);
            this.cmdCancel.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.Size = new System.Drawing.Size(100, 28);
            this.cmdCancel.TabIndex = 36;
            this.cmdCancel.Text = "Cancel";
            this.cmdCancel.UseVisualStyleBackColor = true;
            this.cmdCancel.Click += new System.EventHandler(this.cmdCancel_Click);
            // 
            // picSubject
            // 
            this.picSubject.Image = ((System.Drawing.Image)(resources.GetObject("picSubject.Image")));
            this.picSubject.Location = new System.Drawing.Point(552, 135);
            this.picSubject.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.picSubject.Name = "picSubject";
            this.picSubject.Size = new System.Drawing.Size(35, 25);
            this.picSubject.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.picSubject.TabIndex = 37;
            this.picSubject.TabStop = false;
            this.picSubject.Click += new System.EventHandler(this.picSubject_Click);
            // 
            // picMessage
            // 
            this.picMessage.Image = ((System.Drawing.Image)(resources.GetObject("picMessage.Image")));
            this.picMessage.Location = new System.Drawing.Point(552, 172);
            this.picMessage.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.picMessage.Name = "picMessage";
            this.picMessage.Size = new System.Drawing.Size(35, 25);
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
            this.chkDisabled.Location = new System.Drawing.Point(639, 26);
            this.chkDisabled.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.chkDisabled.Name = "chkDisabled";
            this.chkDisabled.Size = new System.Drawing.Size(85, 21);
            this.chkDisabled.TabIndex = 8;
            this.chkDisabled.Text = "Disabled";
            this.chkDisabled.UseVisualStyleBackColor = true;
            this.chkDisabled.CheckedChanged += new System.EventHandler(this.chkDisabled_CheckedChanged);
            // 
            // webBrowser
            // 
            this.webBrowser.Location = new System.Drawing.Point(107, 172);
            this.webBrowser.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.webBrowser.MinimumSize = new System.Drawing.Size(27, 25);
            this.webBrowser.Name = "webBrowser";
            this.webBrowser.Size = new System.Drawing.Size(449, 165);
            this.webBrowser.TabIndex = 5;
            // 
            // lblSender
            // 
            this.lblSender.AutoSize = true;
            this.lblSender.Location = new System.Drawing.Point(32, 55);
            this.lblSender.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblSender.Name = "lblSender";
            this.lblSender.Size = new System.Drawing.Size(58, 17);
            this.lblSender.TabIndex = 41;
            this.lblSender.Text = "Sender:";
            this.lblSender.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // txtSender
            // 
            this.txtSender.Location = new System.Drawing.Point(107, 51);
            this.txtSender.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtSender.Name = "txtSender";
            this.txtSender.Size = new System.Drawing.Size(448, 22);
            this.txtSender.TabIndex = 1;
            this.txtSender.TextChanged += new System.EventHandler(this.txtSender_TextChanged);
            // 
            // txtCC
            // 
            this.txtCC.Location = new System.Drawing.Point(107, 106);
            this.txtCC.Margin = new System.Windows.Forms.Padding(4);
            this.txtCC.Name = "txtCC";
            this.txtCC.Size = new System.Drawing.Size(448, 22);
            this.txtCC.TabIndex = 3;
            this.txtCC.TextChanged += new System.EventHandler(this.txtCC_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(60, 109);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(30, 17);
            this.label2.TabIndex = 43;
            this.label2.Text = "CC:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // AutoEventDetails
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(975, 731);
            this.Controls.Add(this.txtCC);
            this.Controls.Add(this.label2);
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
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
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
        private System.Windows.Forms.TextBox txtCC;
        private System.Windows.Forms.Label label2;
    }
}