using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace VSIconSwitcher
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            Status = null;
        }

        public event EventHandler<EventArgs> PatchButtonClicked;
        public event EventHandler<EventArgs> UndoButtonClicked;
        public event EventHandler<EventArgs> QuitButtonClicked;
        public event EventHandler<EventArgs> VS10PathChanged;
        public event EventHandler<EventArgs> VS11PathChanged;

        public string VS10Path { get { return m_vs10PathTextBox.Text; } set { m_vs10PathTextBox.Text = value; } }
        public string VS10DetectedProduct { set { m_vs10DetectedProductTextbox.Text = value; } }
        public string VS10Languages { set { m_vs10InstalledLanguagesTextbox.Text = value; } }
        public string VS11Path { get { return m_vs11PathTextBox.Text; } set { m_vs11PathTextBox.Text = value; } }
        public string VS11DetectedProduct { set { m_vs11DetectedProductTextbox.Text = value; } }
        public string VS11Languages { set { m_vs11InstalledLanguagesTextbox.Text = value; } }
        public string BackupPath { get { return m_backupPathTextBox.Text; } set { m_backupPathTextBox.Text = value; } }
        public string Status { set { m_statusTextbox.Text = string.IsNullOrEmpty(value)? "Ready" : value; } }
        public int CurrentProgress { get { return m_progressBar.Value; } set { m_progressBar.Value = value; } }
        public int ProgressMax { set { m_progressBar.Maximum = value; } }

        public bool IsBusy
        {
            set
            {
                m_vs10BrowseButton.Enabled = !value;
                m_vs10PathTextBox.Enabled = !value;
                m_vs11BrowseButton.Enabled = !value;
                m_vs11PathTextBox.Enabled = !value;
                m_backupPathTextBox.Enabled = !value;
                m_patchButton.Enabled = !value;
                m_undoBotton.Enabled = !value;
                //m_quitButton.Enabled = !value;
            }
        }

        private void m_patchButton_Click(object sender, EventArgs e)
        {
            PatchButtonClicked(this, e);
        }

        private void m_undoBotton_Click(object sender, EventArgs e)
        {
            UndoButtonClicked(this, e);
        }

        private void m_vs10BrowseButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.ShowNewFolderButton = false;
            if (Directory.Exists(m_vs10PathTextBox.Text))
            {
                fbd.SelectedPath = m_vs10PathTextBox.Text;
            }
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                m_vs10PathTextBox.Text = fbd.SelectedPath;
            }
        }

        private void m_vs11BrowseButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.ShowNewFolderButton = false;
            if (Directory.Exists(m_vs11PathTextBox.Text))
            {
                fbd.SelectedPath = m_vs11PathTextBox.Text;
            }
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                m_vs11PathTextBox.Text = fbd.SelectedPath;
            }
        }

        private void m_quitButton_Click(object sender, EventArgs e)
        {
            QuitButtonClicked(this, e);
        }

        private void m_vs10PathTextBox_TextChanged(object sender, EventArgs e)
        {
            VS10PathChanged(this, e);
        }

        private void m_vs11PathTextBox_TextChanged(object sender, EventArgs e)
        {
            VS11PathChanged(this, e);
        }
    }
}
