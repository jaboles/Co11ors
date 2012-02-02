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

        public string VS10Path { get { return m_vs10PathTextBox.Text; } set { m_vs10PathTextBox.Text = value; } }
        public string VS11Path { get { return m_vs11PathTextBox.Text; } set { m_vs11PathTextBox.Text = value; } }
        public string BackupPath { get { return m_backupPathTextBox.Text; } set { m_backupPathTextBox.Text = value; } }
        public string Status { set { m_statusTextbox.Text = string.IsNullOrEmpty(value)? "Ready" : value; } }
        public int CurrentProgress { get { return m_progressBar.Value; } set { m_progressBar.Value = value; } }
        public int ProgressMax { set { m_progressBar.Maximum = value; } }

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
    }
}
