using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace InfoQuest.forms
{
    public partial class LoadInfoForm : Form
    {
        public LoadInfoForm()
        {
            InitializeComponent();
            PopulateTreeView();
        }
        private void PopulateTreeView()
        {
            TreeNode rootNode;
            DriveInfo[] allDrives = DriveInfo.GetDrives();
            for (int i = 0; i < allDrives.Length; i++)
            {
                DriveInfo d = allDrives[i];
                if (d.IsReady)
                {
                    //add it to the tree view
                    rootNode = new TreeNode(d.Name);
                    rootNode.Tag = "drive";
                    treeView1.Nodes.Add(rootNode);
                }
            }
            //DirectoryInfo info = new DirectoryInfo(@"D:\");
            //if (info.Exists)
            //{
            //    rootNode = new TreeNode(info.Name);
            //    rootNode.Tag = info;
            //    GetDirectories(info.GetDirectories(), rootNode);
            //    treeView1.Nodes.Add(rootNode);
            //}
        }

        private void GetDirectories(DirectoryInfo[] subDirs,
                                        TreeNode nodeToAddTo)
        {
            try
            {

                TreeNode aNode;
                DirectoryInfo[] subSubDirs;
                foreach (DirectoryInfo subDir in subDirs)
                {
                    aNode = new TreeNode(subDir.Name, 0, 0);
                    aNode.Tag = subDir;
                    aNode.ImageKey = "folder";
                    subSubDirs = subDir.GetDirectories();
                    //if (subSubDirs.Length != 0)
                    //{
                    //    GetDirectories(subSubDirs, aNode);
                    //}
                    nodeToAddTo.Nodes.Add(aNode);
                }
            }
            catch (UnauthorizedAccessException e)
            {
                //igonore
            }
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            TreeNode newSelected = this.treeView1.SelectedNode;
            listView1.Items.Clear();
            DirectoryInfo nodeDirInfo  = null;
            if (newSelected.Tag is String)
               nodeDirInfo  = new DirectoryInfo(newSelected.Text);
            else
               nodeDirInfo = (DirectoryInfo)newSelected.Tag;
            ListViewItem.ListViewSubItem[] subItems;
            ListViewItem item = null;

            foreach (DirectoryInfo dir in nodeDirInfo.GetDirectories())
            {
                item = new ListViewItem(dir.Name, 0);
                subItems = new ListViewItem.ListViewSubItem[]
            {new ListViewItem.ListViewSubItem(item, "Directory"), 
             new ListViewItem.ListViewSubItem(item, 
                dir.LastAccessTime.ToShortDateString())};
                item.SubItems.AddRange(subItems);
                listView1.Items.Add(item);
            }
            foreach (FileInfo file in nodeDirInfo.GetFiles())
            {
                item = new ListViewItem(file.Name, 1);
                subItems = new ListViewItem.ListViewSubItem[]
                            {   new ListViewItem.ListViewSubItem(item, "File"), 
                                new ListViewItem.ListViewSubItem(item, file.LastAccessTime.ToShortDateString())};

                item.SubItems.AddRange(subItems);
                listView1.Items.Add(item);
            }

            listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
        }

        private void listView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            this.textBox1.Text = sender.ToString();

            //this.textBox1.Text = this.treeView1.SelectedNode.Text + this.listView1.SelectedItems[0].Text;
        }

    }
}