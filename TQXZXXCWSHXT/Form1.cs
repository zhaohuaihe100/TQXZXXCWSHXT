using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TQXZXXCWSHXT
{
    public partial class Form1 : Form
    {
        private string selectedSchool;
        private string selectedK1;//单击选中的总分类科目
        private string selectedK2;//单击选中的明细分类科目
        public Form1()
        {
            InitializeComponent();
        }

        private void listBox1_Click(object sender, EventArgs e)
        {
            this.selectedSchool = this.listBox1.SelectedItem.ToString().Substring(2);
            this.lb_welcome.Text = "欢迎" + this.selectedSchool + "前来报账！";
            this.textBbzid.Text = DateTime.Now.ToString();

        }

        private void treeView1_MouseDown(object sender, MouseEventArgs e)
        {
            if ((sender as TreeView) != null)
            {
                treeView1.SelectedNode = treeView1.GetNodeAt(e.X, e.Y);
            }
            
        }

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (this.treeView1.SelectedNode.Text != null && this.treeView1.SelectedNode.Parent != null)
                this.selectedK2 = this.treeView1.SelectedNode.Text;
            if (this.treeView1.SelectedNode.Parent != null)
            {
                this.selectedK1 = this.treeView1.SelectedNode.Parent.Text;
                //MessageBox.Show(this.selectedK1);
            }
            this.lb_selectedK2.Text = this.selectedK2;
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int index = this.dataGridView1.Rows.Add();
            this.dataGridView1.Rows[index].Cells[0].Value = this.selectedSchool;
            this.dataGridView1.Rows[index].Cells[1].Value = this.textBbzid.Text;
            this.dataGridView1.Rows[index].Cells[2].Value = this.selectedK1;
            this.dataGridView1.Rows[index].Cells[3].Value = this.selectedK2;
            this.dataGridView1.Rows[index].Cells[4].Value = this.textBzs.Text;
            this.dataGridView1.Rows[index].Cells[5].Value = this.textBje.Text;
            if (this.checkBsfxj.Checked == true)
                this.dataGridView1.Rows[index].Cells[6].Value = "现金";
            else
                this.dataGridView1.Rows[index].Cells[6].Value = "转账";
            if (this.checkBsfhg.Checked == true)
                this.dataGridView1.Rows[index].Cells[7].Value = "合格";
            else
                this.dataGridView1.Rows[index].Cells[7].Value = "不合格";

            this.dataGridView1.Rows[index].Cells[8].Value = this.textBbz.Text;
            
        }

        //private void treeView1_Click(object sender, EventArgs e)
        //{
            

        //}

        //private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        //{
            
        //}
    }
}
