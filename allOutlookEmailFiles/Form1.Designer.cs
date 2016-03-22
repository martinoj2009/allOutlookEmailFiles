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
            this.registryUsers = new System.Windows.Forms.ListBox();
            this.allOutlookFiles = new System.Windows.Forms.ListBox();
            this.GETFILES = new System.ComponentModel.BackgroundWorker();
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
            this.allOutlookFiles.Size = new System.Drawing.Size(299, 316);
            this.allOutlookFiles.TabIndex = 1;
            this.allOutlookFiles.SelectedIndexChanged += new System.EventHandler(this.allOutlookFiles_SelectedIndexChanged);
            // 
            // GETFILES
            // 
            this.GETFILES.DoWork += new System.ComponentModel.DoWorkEventHandler(this.GETFILES_DoWork);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(551, 378);
            this.Controls.Add(this.allOutlookFiles);
            this.Controls.Add(this.registryUsers);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox registryUsers;
        private System.Windows.Forms.ListBox allOutlookFiles;
        private System.ComponentModel.BackgroundWorker GETFILES;
    }
}

