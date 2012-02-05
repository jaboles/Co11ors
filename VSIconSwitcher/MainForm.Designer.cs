namespace VSIconSwitcher
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
            this.m_descriptionLabel = new System.Windows.Forms.Label();
            this.m_strongNamLabel = new System.Windows.Forms.Label();
            this.m_vs2010Label = new System.Windows.Forms.Label();
            this.m_vs10PathTextBox = new System.Windows.Forms.TextBox();
            this.m_vs11Label = new System.Windows.Forms.Label();
            this.m_vs11PathTextBox = new System.Windows.Forms.TextBox();
            this.m_patchButton = new System.Windows.Forms.Button();
            this.m_undoBotton = new System.Windows.Forms.Button();
            this.m_backupPathLabel = new System.Windows.Forms.Label();
            this.m_backupPathTextBox = new System.Windows.Forms.TextBox();
            this.m_vs10BrowseButton = new System.Windows.Forms.Button();
            this.m_vs11BrowseButton = new System.Windows.Forms.Button();
            this.m_quitButton = new System.Windows.Forms.Button();
            this.m_progressBar = new System.Windows.Forms.ProgressBar();
            this.m_statusLabel = new System.Windows.Forms.Label();
            this.m_statusTextbox = new System.Windows.Forms.TextBox();
            this.m_vs10DetectedProductLabel = new System.Windows.Forms.Label();
            this.m_vs10DetectedProductTextbox = new System.Windows.Forms.TextBox();
            this.m_vs11DetectedProductTextbox = new System.Windows.Forms.TextBox();
            this.m_vs11DetectedProductLabel = new System.Windows.Forms.Label();
            this.m_vs10LanguagesLabel = new System.Windows.Forms.Label();
            this.m_vs10InstalledLanguagesTextbox = new System.Windows.Forms.TextBox();
            this.m_vs11InstalledLanguagesTextbox = new System.Windows.Forms.TextBox();
            this.m_vs11InstalledLanguagesLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // m_descriptionLabel
            // 
            this.m_descriptionLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_descriptionLabel.Location = new System.Drawing.Point(13, 13);
            this.m_descriptionLabel.Name = "m_descriptionLabel";
            this.m_descriptionLabel.Size = new System.Drawing.Size(553, 41);
            this.m_descriptionLabel.TabIndex = 0;
            this.m_descriptionLabel.Text = "This tool allows you to replace the black and white icons in VS 11 with the more " +
    "user-friendly, colourful icons in VS 2010.";
            // 
            // m_strongNamLabel
            // 
            this.m_strongNamLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_strongNamLabel.Location = new System.Drawing.Point(13, 54);
            this.m_strongNamLabel.Name = "m_strongNamLabel";
            this.m_strongNamLabel.Size = new System.Drawing.Size(554, 42);
            this.m_strongNamLabel.TabIndex = 1;
            this.m_strongNamLabel.Text = "Note: modifying managed assemblies will require bypassing strong-name verificatio" +
    "n for those assemblies.";
            // 
            // m_vs2010Label
            // 
            this.m_vs2010Label.Location = new System.Drawing.Point(13, 110);
            this.m_vs2010Label.Name = "m_vs2010Label";
            this.m_vs2010Label.Size = new System.Drawing.Size(187, 18);
            this.m_vs2010Label.TabIndex = 2;
            this.m_vs2010Label.Text = "VS 2010 installation or media path:";
            this.m_vs2010Label.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // m_vs10PathTextBox
            // 
            this.m_vs10PathTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_vs10PathTextBox.Location = new System.Drawing.Point(206, 108);
            this.m_vs10PathTextBox.Name = "m_vs10PathTextBox";
            this.m_vs10PathTextBox.Size = new System.Drawing.Size(333, 20);
            this.m_vs10PathTextBox.TabIndex = 3;
            this.m_vs10PathTextBox.TextChanged += new System.EventHandler(this.m_vs10PathTextBox_TextChanged);
            // 
            // m_vs11Label
            // 
            this.m_vs11Label.Location = new System.Drawing.Point(12, 200);
            this.m_vs11Label.Name = "m_vs11Label";
            this.m_vs11Label.Size = new System.Drawing.Size(188, 20);
            this.m_vs11Label.TabIndex = 4;
            this.m_vs11Label.Text = "VS 11 installation or media path:";
            this.m_vs11Label.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // m_vs11PathTextBox
            // 
            this.m_vs11PathTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_vs11PathTextBox.Location = new System.Drawing.Point(206, 200);
            this.m_vs11PathTextBox.Name = "m_vs11PathTextBox";
            this.m_vs11PathTextBox.Size = new System.Drawing.Size(333, 20);
            this.m_vs11PathTextBox.TabIndex = 5;
            this.m_vs11PathTextBox.TextChanged += new System.EventHandler(this.m_vs11PathTextBox_TextChanged);
            // 
            // m_patchButton
            // 
            this.m_patchButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.m_patchButton.Location = new System.Drawing.Point(210, 395);
            this.m_patchButton.Name = "m_patchButton";
            this.m_patchButton.Size = new System.Drawing.Size(143, 23);
            this.m_patchButton.TabIndex = 6;
            this.m_patchButton.Text = "Backup && Patch";
            this.m_patchButton.UseVisualStyleBackColor = true;
            this.m_patchButton.Click += new System.EventHandler(this.m_patchButton_Click);
            // 
            // m_undoBotton
            // 
            this.m_undoBotton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.m_undoBotton.Location = new System.Drawing.Point(359, 395);
            this.m_undoBotton.Name = "m_undoBotton";
            this.m_undoBotton.Size = new System.Drawing.Size(131, 23);
            this.m_undoBotton.TabIndex = 7;
            this.m_undoBotton.Text = "Undo";
            this.m_undoBotton.UseVisualStyleBackColor = true;
            this.m_undoBotton.Click += new System.EventHandler(this.m_undoBotton_Click);
            // 
            // m_backupPathLabel
            // 
            this.m_backupPathLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.m_backupPathLabel.Location = new System.Drawing.Point(12, 308);
            this.m_backupPathLabel.Name = "m_backupPathLabel";
            this.m_backupPathLabel.Size = new System.Drawing.Size(186, 18);
            this.m_backupPathLabel.TabIndex = 8;
            this.m_backupPathLabel.Text = "Folder for backups";
            this.m_backupPathLabel.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // m_backupPathTextBox
            // 
            this.m_backupPathTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_backupPathTextBox.Location = new System.Drawing.Point(206, 308);
            this.m_backupPathTextBox.Name = "m_backupPathTextBox";
            this.m_backupPathTextBox.Size = new System.Drawing.Size(360, 20);
            this.m_backupPathTextBox.TabIndex = 9;
            // 
            // m_vs10BrowseButton
            // 
            this.m_vs10BrowseButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_vs10BrowseButton.Location = new System.Drawing.Point(544, 108);
            this.m_vs10BrowseButton.Name = "m_vs10BrowseButton";
            this.m_vs10BrowseButton.Size = new System.Drawing.Size(28, 23);
            this.m_vs10BrowseButton.TabIndex = 10;
            this.m_vs10BrowseButton.Text = "...";
            this.m_vs10BrowseButton.UseVisualStyleBackColor = true;
            this.m_vs10BrowseButton.Click += new System.EventHandler(this.m_vs10BrowseButton_Click);
            // 
            // m_vs11BrowseButton
            // 
            this.m_vs11BrowseButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.m_vs11BrowseButton.Location = new System.Drawing.Point(544, 198);
            this.m_vs11BrowseButton.Name = "m_vs11BrowseButton";
            this.m_vs11BrowseButton.Size = new System.Drawing.Size(28, 23);
            this.m_vs11BrowseButton.TabIndex = 11;
            this.m_vs11BrowseButton.Text = "...";
            this.m_vs11BrowseButton.UseVisualStyleBackColor = true;
            this.m_vs11BrowseButton.Click += new System.EventHandler(this.m_vs11BrowseButton_Click);
            // 
            // m_quitButton
            // 
            this.m_quitButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.m_quitButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.m_quitButton.Location = new System.Drawing.Point(496, 395);
            this.m_quitButton.Name = "m_quitButton";
            this.m_quitButton.Size = new System.Drawing.Size(71, 23);
            this.m_quitButton.TabIndex = 12;
            this.m_quitButton.Text = "Quit";
            this.m_quitButton.UseVisualStyleBackColor = true;
            this.m_quitButton.Click += new System.EventHandler(this.m_quitButton_Click);
            // 
            // m_progressBar
            // 
            this.m_progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_progressBar.Location = new System.Drawing.Point(17, 398);
            this.m_progressBar.Name = "m_progressBar";
            this.m_progressBar.Size = new System.Drawing.Size(166, 15);
            this.m_progressBar.TabIndex = 13;
            // 
            // m_statusLabel
            // 
            this.m_statusLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.m_statusLabel.AutoSize = true;
            this.m_statusLabel.Location = new System.Drawing.Point(14, 343);
            this.m_statusLabel.Name = "m_statusLabel";
            this.m_statusLabel.Size = new System.Drawing.Size(37, 13);
            this.m_statusLabel.TabIndex = 14;
            this.m_statusLabel.Text = "Status";
            // 
            // m_statusTextbox
            // 
            this.m_statusTextbox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_statusTextbox.Location = new System.Drawing.Point(17, 359);
            this.m_statusTextbox.Name = "m_statusTextbox";
            this.m_statusTextbox.ReadOnly = true;
            this.m_statusTextbox.Size = new System.Drawing.Size(546, 20);
            this.m_statusTextbox.TabIndex = 15;
            // 
            // m_vs10DetectedProductLabel
            // 
            this.m_vs10DetectedProductLabel.Location = new System.Drawing.Point(12, 134);
            this.m_vs10DetectedProductLabel.Name = "m_vs10DetectedProductLabel";
            this.m_vs10DetectedProductLabel.Size = new System.Drawing.Size(187, 18);
            this.m_vs10DetectedProductLabel.TabIndex = 16;
            this.m_vs10DetectedProductLabel.Text = "Detected product:";
            this.m_vs10DetectedProductLabel.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // m_vs10DetectedProductTextbox
            // 
            this.m_vs10DetectedProductTextbox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_vs10DetectedProductTextbox.Location = new System.Drawing.Point(206, 134);
            this.m_vs10DetectedProductTextbox.Name = "m_vs10DetectedProductTextbox";
            this.m_vs10DetectedProductTextbox.ReadOnly = true;
            this.m_vs10DetectedProductTextbox.Size = new System.Drawing.Size(244, 20);
            this.m_vs10DetectedProductTextbox.TabIndex = 17;
            // 
            // m_vs11DetectedProductTextbox
            // 
            this.m_vs11DetectedProductTextbox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_vs11DetectedProductTextbox.Location = new System.Drawing.Point(206, 226);
            this.m_vs11DetectedProductTextbox.Name = "m_vs11DetectedProductTextbox";
            this.m_vs11DetectedProductTextbox.ReadOnly = true;
            this.m_vs11DetectedProductTextbox.Size = new System.Drawing.Size(244, 20);
            this.m_vs11DetectedProductTextbox.TabIndex = 18;
            // 
            // m_vs11DetectedProductLabel
            // 
            this.m_vs11DetectedProductLabel.Location = new System.Drawing.Point(13, 226);
            this.m_vs11DetectedProductLabel.Name = "m_vs11DetectedProductLabel";
            this.m_vs11DetectedProductLabel.Size = new System.Drawing.Size(187, 18);
            this.m_vs11DetectedProductLabel.TabIndex = 19;
            this.m_vs11DetectedProductLabel.Text = "Detected product:";
            this.m_vs11DetectedProductLabel.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // m_vs10LanguagesLabel
            // 
            this.m_vs10LanguagesLabel.Location = new System.Drawing.Point(13, 160);
            this.m_vs10LanguagesLabel.Name = "m_vs10LanguagesLabel";
            this.m_vs10LanguagesLabel.Size = new System.Drawing.Size(187, 18);
            this.m_vs10LanguagesLabel.TabIndex = 20;
            this.m_vs10LanguagesLabel.Text = "Installed languages:";
            this.m_vs10LanguagesLabel.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // m_vs10InstalledLanguagesTextbox
            // 
            this.m_vs10InstalledLanguagesTextbox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_vs10InstalledLanguagesTextbox.Location = new System.Drawing.Point(206, 160);
            this.m_vs10InstalledLanguagesTextbox.Name = "m_vs10InstalledLanguagesTextbox";
            this.m_vs10InstalledLanguagesTextbox.ReadOnly = true;
            this.m_vs10InstalledLanguagesTextbox.Size = new System.Drawing.Size(244, 20);
            this.m_vs10InstalledLanguagesTextbox.TabIndex = 21;
            // 
            // m_vs11InstalledLanguagesTextbox
            // 
            this.m_vs11InstalledLanguagesTextbox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_vs11InstalledLanguagesTextbox.Location = new System.Drawing.Point(206, 252);
            this.m_vs11InstalledLanguagesTextbox.Name = "m_vs11InstalledLanguagesTextbox";
            this.m_vs11InstalledLanguagesTextbox.ReadOnly = true;
            this.m_vs11InstalledLanguagesTextbox.Size = new System.Drawing.Size(244, 20);
            this.m_vs11InstalledLanguagesTextbox.TabIndex = 22;
            // 
            // m_vs11InstalledLanguagesLabel
            // 
            this.m_vs11InstalledLanguagesLabel.Location = new System.Drawing.Point(13, 252);
            this.m_vs11InstalledLanguagesLabel.Name = "m_vs11InstalledLanguagesLabel";
            this.m_vs11InstalledLanguagesLabel.Size = new System.Drawing.Size(187, 18);
            this.m_vs11InstalledLanguagesLabel.TabIndex = 23;
            this.m_vs11InstalledLanguagesLabel.Text = "Installed languages:";
            this.m_vs11InstalledLanguagesLabel.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.m_quitButton;
            this.ClientSize = new System.Drawing.Size(578, 430);
            this.Controls.Add(this.m_vs11InstalledLanguagesLabel);
            this.Controls.Add(this.m_vs11InstalledLanguagesTextbox);
            this.Controls.Add(this.m_vs10InstalledLanguagesTextbox);
            this.Controls.Add(this.m_vs10LanguagesLabel);
            this.Controls.Add(this.m_vs11DetectedProductLabel);
            this.Controls.Add(this.m_vs11DetectedProductTextbox);
            this.Controls.Add(this.m_vs10DetectedProductTextbox);
            this.Controls.Add(this.m_vs10DetectedProductLabel);
            this.Controls.Add(this.m_statusTextbox);
            this.Controls.Add(this.m_statusLabel);
            this.Controls.Add(this.m_progressBar);
            this.Controls.Add(this.m_quitButton);
            this.Controls.Add(this.m_vs11BrowseButton);
            this.Controls.Add(this.m_vs10BrowseButton);
            this.Controls.Add(this.m_backupPathTextBox);
            this.Controls.Add(this.m_backupPathLabel);
            this.Controls.Add(this.m_undoBotton);
            this.Controls.Add(this.m_patchButton);
            this.Controls.Add(this.m_vs11PathTextBox);
            this.Controls.Add(this.m_vs11Label);
            this.Controls.Add(this.m_vs10PathTextBox);
            this.Controls.Add(this.m_vs2010Label);
            this.Controls.Add(this.m_strongNamLabel);
            this.Controls.Add(this.m_descriptionLabel);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Co11ors";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label m_descriptionLabel;
        private System.Windows.Forms.Label m_strongNamLabel;
        private System.Windows.Forms.Label m_vs2010Label;
        private System.Windows.Forms.TextBox m_vs10PathTextBox;
        private System.Windows.Forms.Label m_vs11Label;
        private System.Windows.Forms.TextBox m_vs11PathTextBox;
        private System.Windows.Forms.Button m_patchButton;
        private System.Windows.Forms.Button m_undoBotton;
        private System.Windows.Forms.Label m_backupPathLabel;
        private System.Windows.Forms.TextBox m_backupPathTextBox;
        private System.Windows.Forms.Button m_vs10BrowseButton;
        private System.Windows.Forms.Button m_vs11BrowseButton;
        private System.Windows.Forms.Button m_quitButton;
        private System.Windows.Forms.ProgressBar m_progressBar;
        private System.Windows.Forms.Label m_statusLabel;
        private System.Windows.Forms.TextBox m_statusTextbox;
        private System.Windows.Forms.Label m_vs10DetectedProductLabel;
        private System.Windows.Forms.TextBox m_vs10DetectedProductTextbox;
        private System.Windows.Forms.TextBox m_vs11DetectedProductTextbox;
        private System.Windows.Forms.Label m_vs11DetectedProductLabel;
        private System.Windows.Forms.Label m_vs10LanguagesLabel;
        private System.Windows.Forms.TextBox m_vs10InstalledLanguagesTextbox;
        private System.Windows.Forms.TextBox m_vs11InstalledLanguagesTextbox;
        private System.Windows.Forms.Label m_vs11InstalledLanguagesLabel;
    }
}

