namespace allOutlookEmailFiles
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.registryUsers = new System.Windows.Forms.ListBox();
            this.allOutlookFiles = new System.Windows.Forms.ListBox();
            this.GETFILES = new System.ComponentModel.BackgroundWorker();
            this.usernameLabel = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.officeVersionLabel = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // registryUsers
            // 
            this.registryUsers.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.registryUsers.FormattingEnabled = true;
            this.registryUsers.Location = new System.Drawing.Point(12, 50);
            this.registryUsers.Name = "registryUsers";
            this.registryUsers.Size = new System.Drawing.Size(222, 316);
            this.registryUsers.TabIndex = 0;
            this.registryUsers.SelectedIndexChanged += new System.EventHandler(this.registryUsers_SelectedIndexChanged);
            // 
            // allOutlookFiles
            // 
            this.allOutlookFiles.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.allOutlookFiles.FormattingEnabled = true;
            this.allOutlookFiles.HorizontalScrollbar = true;
            this.allOutlookFiles.Location = new System.Drawing.Point(240, 50);
            this.allOutlookFiles.Name = "allOutlookFiles";
            this.allOutlookFiles.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.allOutlookFiles.Size = new System.Drawing.Size(299, 316);
            this.allOutlookFiles.TabIndex = 1;
            this.allOutlookFiles.SelectedIndexChanged += new System.EventHandler(this.allOutlookFiles_SelectedIndexChanged);
            // 
            // GETFILES
            // 
            this.GETFILES.DoWork += new System.ComponentModel.DoWorkEventHandler(this.GETFILES_DoWork);
            // 
            // usernameLabel
            // 
            this.usernameLabel.AutoSize = true;
            this.usernameLabel.Location = new System.Drawing.Point(13, 31);
            this.usernameLabel.Name = "usernameLabel";
            this.usernameLabel.Size = new System.Drawing.Size(68, 13);
            this.usernameLabel.TabIndex = 2;
            this.usernameLabel.Text = "USERNAME";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(240, 31);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(79, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Office Version: ";
            // 
            // officeVersionLabel
            // 
            this.officeVersionLabel.AutoSize = true;
            this.officeVersionLabel.Location = new System.Drawing.Point(326, 31);
            this.officeVersionLabel.Name = "officeVersionLabel";
            this.officeVersionLabel.Size = new System.Drawing.Size(25, 13);
            this.officeVersionLabel.TabIndex = 4;
            this.officeVersionLabel.Text = "000";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(551, 24);
            this.menuStrip1.TabIndex = 5;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "Help";
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(107, 22);
            this.aboutToolStripMenuItem.Text = "About";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(551, 378);
            this.Controls.Add(this.officeVersionLabel);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.usernameLabel);
            this.Controls.Add(this.allOutlookFiles);
            this.Controls.Add(this.registryUsers);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Get Outlook Files";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox registryUsers;
        private System.Windows.Forms.ListBox allOutlookFiles;
        private System.ComponentModel.BackgroundWorker GETFILES;
        private System.Windows.Forms.Label usernameLabel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label officeVersionLabel;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
    }
}

