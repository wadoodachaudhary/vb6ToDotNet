namespace AutoNotice
{
    partial class AutoNoticeEditor
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AutoNoticeEditor));
            this.toolStripContainer1 = new System.Windows.Forms.ToolStripContainer();
            this.picAutoEvents = new System.Windows.Forms.PictureBox();
            this.picAutoReports = new System.Windows.Forms.PictureBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.statusBar1 = new System.Windows.Forms.StatusBar();
            this.mainMenu = new System.Windows.Forms.MenuStrip();
            this.mnuNotifications = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuExit = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuHelp = new System.Windows.Forms.ToolStripMenuItem();
            this.helpItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripContainer1.ContentPanel.SuspendLayout();
            this.toolStripContainer1.TopToolStripPanel.SuspendLayout();
            this.toolStripContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picAutoEvents)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picAutoReports)).BeginInit();
            this.mainMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStripContainer1
            // 
            // 
            // toolStripContainer1.ContentPanel
            // 
            this.toolStripContainer1.ContentPanel.Controls.Add(this.picAutoEvents);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.picAutoReports);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.label2);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.label1);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.statusBar1);
            this.toolStripContainer1.ContentPanel.Size = new System.Drawing.Size(745, 322);
            this.toolStripContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.toolStripContainer1.Location = new System.Drawing.Point(0, 0);
            this.toolStripContainer1.Name = "toolStripContainer1";
            this.toolStripContainer1.Size = new System.Drawing.Size(745, 346);
            this.toolStripContainer1.TabIndex = 0;
            this.toolStripContainer1.Text = "toolStripContainer1";
            // 
            // toolStripContainer1.TopToolStripPanel
            // 
            this.toolStripContainer1.TopToolStripPanel.Controls.Add(this.mainMenu);
            // 
            // picAutoEvents
            // 
            this.picAutoEvents.Image = ((System.Drawing.Image)(resources.GetObject("picAutoEvents.Image")));
            this.picAutoEvents.Location = new System.Drawing.Point(434, 83);
            this.picAutoEvents.Name = "picAutoEvents";
            this.picAutoEvents.Size = new System.Drawing.Size(194, 159);
            this.picAutoEvents.TabIndex = 5;
            this.picAutoEvents.TabStop = false;
            this.picAutoEvents.MouseLeave += new System.EventHandler(this.picAutoEvents_MouseLeave);
            this.picAutoEvents.Click += new System.EventHandler(this.picAutoEvents_Click);
            this.picAutoEvents.MouseHover += new System.EventHandler(this.picAutoEvents_MouseHover);
            // 
            // picAutoReports
            // 
            this.picAutoReports.Image = ((System.Drawing.Image)(resources.GetObject("picAutoReports.Image")));
            this.picAutoReports.Location = new System.Drawing.Point(66, 83);
            this.picAutoReports.Name = "picAutoReports";
            this.picAutoReports.Size = new System.Drawing.Size(243, 159);
            this.picAutoReports.TabIndex = 4;
            this.picAutoReports.TabStop = false;
            this.picAutoReports.MouseLeave += new System.EventHandler(this.picAutoReports_MouseLeave);
            this.picAutoReports.Click += new System.EventHandler(this.picAutoReports_Click);
            this.picAutoReports.MouseHover += new System.EventHandler(this.picAutoReports_MouseHover);
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Arial", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(443, 46);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(195, 29);
            this.label2.TabIndex = 3;
            this.label2.Text = "Auto Notifications";
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Arial", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(62, 46);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(293, 34);
            this.label1.TabIndex = 1;
            this.label1.Text = "Campaign and Event Manager";
            // 
            // statusBar1
            // 
            this.statusBar1.Location = new System.Drawing.Point(0, 300);
            this.statusBar1.Name = "statusBar1";
            this.statusBar1.Size = new System.Drawing.Size(745, 22);
            this.statusBar1.TabIndex = 0;
            this.statusBar1.Text = "Click icon to lanch a task.   Alt+ X - Exit";
            // 
            // mainMenu
            // 
            this.mainMenu.Dock = System.Windows.Forms.DockStyle.None;
            this.mainMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuNotifications,
            this.mnuHelp});
            this.mainMenu.Location = new System.Drawing.Point(0, 0);
            this.mainMenu.Name = "mainMenu";
            this.mainMenu.Size = new System.Drawing.Size(745, 24);
            this.mainMenu.TabIndex = 0;
            this.mainMenu.Text = "menuStrip1";
            // 
            // mnuNotifications
            // 
            this.mnuNotifications.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuExit});
            this.mnuNotifications.Name = "mnuNotifications";
            this.mnuNotifications.Size = new System.Drawing.Size(78, 20);
            this.mnuNotifications.Text = "&Notifications";
            this.mnuNotifications.Visible = false;
            // 
            // mnuExit
            // 
            this.mnuExit.Name = "mnuExit";
            this.mnuExit.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.X)));
            this.mnuExit.Size = new System.Drawing.Size(137, 22);
            this.mnuExit.Text = "Exit";
            this.mnuExit.Click += new System.EventHandler(this.mnuExit_Click);
            // 
            // mnuHelp
            // 
            this.mnuHelp.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.helpItem,
            this.aboutItem});
            this.mnuHelp.Name = "mnuHelp";
            this.mnuHelp.Size = new System.Drawing.Size(40, 20);
            this.mnuHelp.Text = "&Help";
            // 
            // helpItem
            // 
            this.helpItem.Name = "helpItem";
            this.helpItem.ShortcutKeys = System.Windows.Forms.Keys.F1;
            this.helpItem.Size = new System.Drawing.Size(197, 22);
            this.helpItem.Text = "Help";
            // 
            // aboutItem
            // 
            this.aboutItem.Name = "aboutItem";
            this.aboutItem.Size = new System.Drawing.Size(197, 22);
            this.aboutItem.Text = "About Workflow Builder";
            this.aboutItem.Click += new System.EventHandler(this.aboutItem_Click);
            // 
            // AutoNoticeEditor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(745, 346);
            this.Controls.Add(this.toolStripContainer1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.mainMenu;
            this.Name = "AutoNoticeEditor";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Workflow Builder";
            this.Load += new System.EventHandler(this.AutoNoticeEditor_Load);
            this.toolStripContainer1.ContentPanel.ResumeLayout(false);
            this.toolStripContainer1.TopToolStripPanel.ResumeLayout(false);
            this.toolStripContainer1.TopToolStripPanel.PerformLayout();
            this.toolStripContainer1.ResumeLayout(false);
            this.toolStripContainer1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picAutoEvents)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picAutoReports)).EndInit();
            this.mainMenu.ResumeLayout(false);
            this.mainMenu.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ToolStripContainer toolStripContainer1;
        private System.Windows.Forms.MenuStrip mainMenu;
        private System.Windows.Forms.ToolStripMenuItem mnuNotifications;
        private System.Windows.Forms.StatusBar statusBar1;
        private System.Windows.Forms.ToolStripMenuItem mnuExit;
        private System.Windows.Forms.ToolStripMenuItem mnuHelp;
        private System.Windows.Forms.ToolStripMenuItem helpItem;
        private System.Windows.Forms.ToolStripMenuItem aboutItem;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.PictureBox picAutoEvents;
        private System.Windows.Forms.PictureBox picAutoReports;
    }
}

