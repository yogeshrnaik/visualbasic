using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using InfoQuest.forms;

namespace InfoQuest
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void mExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void mLoadInfo_Click(object sender, EventArgs e)
        {
            //LoadInfoForm frm = new LoadInfoForm();
            //frm.WindowState = FormWindowState.Maximized;
            //frm.ShowDialog();
            //frm = null;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            LoadInfoForm frm = new LoadInfoForm();
            frm.WindowState = FormWindowState.Maximized;
            frm.ShowDialog();
            frm = null;
            this.Close();
        }
    }
}