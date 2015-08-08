namespace InfoQuest
{
    partial class MainForm
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
            this.mainMenu = new System.Windows.Forms.MenuStrip();
            this.mFileMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.mLoadInfo = new System.Windows.Forms.ToolStripMenuItem();
            this.mFindInfo = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.mExit = new System.Windows.Forms.ToolStripMenuItem();
            this.mToolsMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.mCustomize = new System.Windows.Forms.ToolStripMenuItem();
            this.mOptions = new System.Windows.Forms.ToolStripMenuItem();
            this.mHelpMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.mContents = new System.Windows.Forms.ToolStripMenuItem();
            this.mAbout = new System.Windows.Forms.ToolStripMenuItem();
            this.statusBar = new System.Windows.Forms.StatusStrip();
            this.mainTabCtrl = new System.Windows.Forms.TabControl();
            this.tpLoadInfo = new System.Windows.Forms.TabPage();
            this.mainMenu.SuspendLayout();
            this.mainTabCtrl.SuspendLayout();
            this.SuspendLayout();
            // 
            // mainMenu
            // 
            this.mainMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mFileMenu,
            this.mToolsMenu,
            this.mHelpMenu});
            this.mainMenu.Location = new System.Drawing.Point(0, 0);
            this.mainMenu.Name = "mainMenu";
            this.mainMenu.Size = new System.Drawing.Size(784, 24);
            this.mainMenu.TabIndex = 0;
            this.mainMenu.Text = "menuStrip1";
            // 
            // mFileMenu
            // 
            this.mFileMenu.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mLoadInfo,
            this.mFindInfo,
            this.toolStripMenuItem1,
            this.mExit});
            this.mFileMenu.Name = "mFileMenu";
            this.mFileMenu.Size = new System.Drawing.Size(40, 20);
            this.mFileMenu.Text = "&File";
            // 
            // mLoadInfo
            // 
            this.mLoadInfo.Name = "mLoadInfo";
            this.mLoadInfo.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.L)));
            this.mLoadInfo.Size = new System.Drawing.Size(184, 22);
            this.mLoadInfo.Text = "&Load Info";
            this.mLoadInfo.Click += new System.EventHandler(this.mLoadInfo_Click);
            // 
            // mFindInfo
            // 
            this.mFindInfo.Name = "mFindInfo";
            this.mFindInfo.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.F)));
            this.mFindInfo.Size = new System.Drawing.Size(184, 22);
            this.mFindInfo.Text = "&Find Info";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(181, 6);
            // 
            // mExit
            // 
            this.mExit.Name = "mExit";
            this.mExit.Size = new System.Drawing.Size(184, 22);
            this.mExit.Text = "E&xit";
            this.mExit.Click += new System.EventHandler(this.mExit_Click);
            // 
            // mToolsMenu
            // 
            this.mToolsMenu.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mCustomize,
            this.mOptions});
            this.mToolsMenu.Name = "mToolsMenu";
            this.mToolsMenu.Size = new System.Drawing.Size(51, 20);
            this.mToolsMenu.Text = "&Tools";
            // 
            // mCustomize
            // 
            this.mCustomize.Name = "mCustomize";
            this.mCustomize.Size = new System.Drawing.Size(148, 22);
            this.mCustomize.Text = "&Customize";
            // 
            // mOptions
            // 
            this.mOptions.Name = "mOptions";
            this.mOptions.Size = new System.Drawing.Size(148, 22);
            this.mOptions.Text = "&Options";
            // 
            // mHelpMenu
            // 
            this.mHelpMenu.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mContents,
            this.mAbout});
            this.mHelpMenu.Name = "mHelpMenu";
            this.mHelpMenu.Size = new System.Drawing.Size(45, 20);
            this.mHelpMenu.Text = "&Help";
            // 
            // mContents
            // 
            this.mContents.Name = "mContents";
            this.mContents.Size = new System.Drawing.Size(139, 22);
            this.mContents.Text = "&Contents";
            // 
            // mAbout
            // 
            this.mAbout.Name = "mAbout";
            this.mAbout.Size = new System.Drawing.Size(139, 22);
            this.mAbout.Text = "&About...";
            // 
            // statusBar
            // 
            this.statusBar.Location = new System.Drawing.Point(0, 519);
            this.statusBar.Name = "statusBar";
            this.statusBar.Size = new System.Drawing.Size(784, 22);
            this.statusBar.TabIndex = 1;
            this.statusBar.Text = "statusStrip1";
            // 
            // mainTabCtrl
            // 
            this.mainTabCtrl.Controls.Add(this.tpLoadInfo);
            this.mainTabCtrl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainTabCtrl.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mainTabCtrl.Location = new System.Drawing.Point(0, 24);
            this.mainTabCtrl.Name = "mainTabCtrl";
            this.mainTabCtrl.SelectedIndex = 0;
            this.mainTabCtrl.Size = new System.Drawing.Size(784, 495);
            this.mainTabCtrl.TabIndex = 2;
            // 
            // tpLoadInfo
            // 
            this.tpLoadInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tpLoadInfo.Location = new System.Drawing.Point(4, 25);
            this.tpLoadInfo.Name = "tpLoadInfo";
            this.tpLoadInfo.Padding = new System.Windows.Forms.Padding(3);
            this.tpLoadInfo.Size = new System.Drawing.Size(776, 466);
            this.tpLoadInfo.TabIndex = 0;
            this.tpLoadInfo.Text = "Load Information";
            this.tpLoadInfo.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 541);
            this.Controls.Add(this.mainTabCtrl);
            this.Controls.Add(this.statusBar);
            this.Controls.Add(this.mainMenu);
            this.MainMenuStrip = this.mainMenu;
            this.Name = "MainForm";
            this.Text = "InfoQuest";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.mainMenu.ResumeLayout(false);
            this.mainMenu.PerformLayout();
            this.mainTabCtrl.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip mainMenu;
        private System.Windows.Forms.StatusStrip statusBar;
        private System.Windows.Forms.TabControl mainTabCtrl;
        private System.Windows.Forms.TabPage tpLoadInfo;
        private System.Windows.Forms.ToolStripMenuItem mFileMenu;
        private System.Windows.Forms.ToolStripMenuItem mLoadInfo;
        private System.Windows.Forms.ToolStripMenuItem mFindInfo;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem mExit;
        private System.Windows.Forms.ToolStripMenuItem mToolsMenu;
        private System.Windows.Forms.ToolStripMenuItem mCustomize;
        private System.Windows.Forms.ToolStripMenuItem mOptions;
        private System.Windows.Forms.ToolStripMenuItem mHelpMenu;
        private System.Windows.Forms.ToolStripMenuItem mContents;
        private System.Windows.Forms.ToolStripMenuItem mAbout;
    }
}

