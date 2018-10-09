using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

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
            this.selectedSchool = "";//初始化选中提示信息
            this.selectedK1 = "";
            this.selectedK2 = "";
        }

        private void listBox1_Click(object sender, EventArgs e)
        {
            this.selectedSchool = this.listBox1.SelectedItem.ToString().Substring(2);
            this.lb_welcome.Text = "欢迎" + this.selectedSchool + "前来报账！";
            //报账ID为年月日+日时分，除年之外，每一项都占2位，
            this.textBbzid.Text = DateTime.Now.ToString("yyyy-MM-dd hh:mm:mm"); ;        // 2008-09-04 

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
            if (this.selectedSchool == "" || this.selectedK1 == "" || this.selectedK2 == "")
                MessageBox.Show("请选择报账学校、报账明细科目");
            else
                {   
                     double tmp;
                     if (this.textBzs.Text =="" || this.textBje.Text ==""||(!double.TryParse(this.textBzs.Text, out tmp))||(!double.TryParse(this.textBje.Text, out tmp)))
                        MessageBox.Show("请输入票据张数及金额！");
                     else
                     {
                         int index = this.dataGridView1.Rows.Add();
                         this.dataGridView1.Rows[index].Cells[0].Value = this.selectedSchool;
                         this.dataGridView1.Rows[index].Cells[1].Value = this.textBbzid.Text;
                         this.dataGridView1.Rows[index].Cells[2].Value = this.selectedK1;
                         this.dataGridView1.Rows[index].Cells[3].Value = this.selectedK2;
                         this.dataGridView1.Rows[index].Cells[6].Value = this.textBzs.Text;//票据张数
                         this.dataGridView1.Rows[index].Cells[7].Value = this.textBje.Text;//票据金额
                         if (this.checkBsfxj.Checked == true)
                             this.dataGridView1.Rows[index].Cells[4].Value = "现金";
                         else
                             this.dataGridView1.Rows[index].Cells[4].Value = "转账";
                         if (this.checkBsfhg.Checked == true)
                             this.dataGridView1.Rows[index].Cells[5].Value = "合格";
                         else
                             this.dataGridView1.Rows[index].Cells[5].Value = "不合格";

                         this.dataGridView1.Rows[index].Cells[8].Value = this.textBbz.Text;


                         this.textBzs.Text = "";
                         this.textBje.Text = "";
                     }
                 }

                
            
        }

        private void label2_Click(object sender, EventArgs e)
        {


        }

        private void button8_Click(object sender, EventArgs e)// 分类汇总
        {
            int nRows=1;
            if (this.dataGridView1.Rows.Count > 1)
                nRows = this.dataGridView1.Rows.Count - 1;
            else
                return;
            
            //MessageBox.Show(nRows.ToString());
            //首先把报账记录原始数据插入到数据库
            //this.InsertJLToMysql();

            //其次，把数据库中的本次报账的数据汇总后返回
            BzjiluHz();
            
            //for (int i = 0; i < nRows; i++)
            //{
            //    RecordHZ tRecordHz = new RecordHZ();
                
                
            //}
        }
        private bool InsertJLToMysql() //把datagrideview中的记录数据插入mysql
        {
            int nRows=1;
            if (this.dataGridView1.Rows.Count > 1)
                nRows = this.dataGridView1.Rows.Count - 1;
            else
                return false;
            for (int i = 0; i < nRows; i++)
            {
                if (this.dataGridView1.Rows[i].Cells[5].Value.ToString() == "小计")
                {
                    nRows = i - 1;//为了防止用户点击了汇总后，又点击保存记录时把汇总项加上。
                }
            }

            MysqlConnector mc = new MysqlConnector();
            mc.SetServer("127.0.0.1");
            mc.SetUserID("cwsh6");
            mc.SetPassword("1234");
            mc.SetDataBase("TQXZXXCWSHXT");
            
            for(int i=0;i<nRows;i++)
            {
                string ssql = "insert into bzjilu values(" + "'" + this.dataGridView1.Rows[i].Cells[0].Value + "'" + "," + "'" + this.dataGridView1.Rows[i].Cells[1].Value + "'" + "," + "'" + this.dataGridView1.Rows[i].Cells[2].Value + "'" + "," + "'" + this.dataGridView1.Rows[i].Cells[3].Value + "'" + "," + this.dataGridView1.Rows[i].Cells[6].Value + "," + this.dataGridView1.Rows[i].Cells[7].Value + "," + "'" + this.dataGridView1.Rows[i].Cells[4].Value + "'" + "," + "'" + this.dataGridView1.Rows[i].Cells[5].Value + "'" + "," + "'" + this.dataGridView1.Rows[i].Cells[8].Value + "'" + ")";
                //MySqlDataAdapter reader = mc.ExeQuery(ssql);
                //MessageBox.Show(ssql);

                //mc.ExeQuery(ssql);
                mc.ExeUpdate(ssql);

            }

            MessageBox.Show("保存完毕");
            return true;
        }

        private bool BzjiluHz()
        {
            MessageBox.Show("请注意，分类汇总后的数据不用再保存到数据库中");
            MysqlConnector mc = new MysqlConnector();
            mc.SetServer("127.0.0.1");
            mc.SetUserID("cwsh6");
            mc.SetPassword("1234");
            mc.SetDataBase("TQXZXXCWSHXT");

            string ssql1="select bzxx,bzid,zflkm,ifnull(mxflkm,'合计：') as mxflkm,ifnull(sfxj,'小计') as sfxj ,";
            string ssql2 = "ifnull(sfhg,'小计') as sfhg,sum(pjzs) as pjzs,sum(pjje) as pjje from bzjilu where bzid='"+this.textBbzid.Text.ToString()+"'"+" group by mxflkm,sfxj,sfhg with rollup"; //from bzjilu where bzid='"+this.textBbzid.Text.ToString()+"'"+groupBox1 
            string ssql = ssql1 + ssql2;
            //MessageBox.Show(ssql);
            //this.dataGridView1.Rows.Clear();
            MySqlDataReader hzjg = mc.ExeQuery(ssql);
            if (!hzjg.Read())
            {
                MessageBox.Show("请先保存数据后再汇总");
                return false;
            }
            while (hzjg.Read())
            {
                int index = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[index].Cells[0].Value = hzjg.GetString(0);//报账学校
                this.dataGridView1.Rows[index].Cells[1].Value = hzjg.GetString(1);//报账ID
                this.dataGridView1.Rows[index].Cells[2].Value = hzjg.GetString(2);//总分类科目if
                this.dataGridView1.Rows[index].Cells[3].Value = hzjg.GetString(3);//明细分类科目
                if (hzjg.GetString(3) == "合计：")
                    this.dataGridView1.Rows[index].Cells[3].Style.ForeColor = Color.Red;
                this.dataGridView1.Rows[index].Cells[4].Value = hzjg.GetString(4);//是否现金
                if (hzjg.GetString(4) == "小计")
                    this.dataGridView1.Rows[index].Cells[4].Style.ForeColor = Color.Red;
                this.dataGridView1.Rows[index].Cells[5].Value = hzjg.GetString(5);//是否合格
                if (hzjg.GetString(5) == "小计")
                    this.dataGridView1.Rows[index].Cells[5].Style.ForeColor = Color.Red;

                this.dataGridView1.Rows[index].Cells[6].Value = hzjg.GetUInt32(6);//票据张数
                this.dataGridView1.Rows[index].Cells[7].Value = hzjg.GetDouble(7);//票据金额
                
            }
            return true;
        }

        private void button9_Click(object sender, EventArgs e)//清空报账记录
        {
            this.dataGridView1.Rows.Clear();

            MysqlConnector mc = new MysqlConnector();
            mc.SetServer("127.0.0.1");
            mc.SetUserID("cwsh6");
            mc.SetPassword("1234");
            mc.SetDataBase("TQXZXXCWSHXT");
            string ssql = "delete from bzjilu where bzid='" + this.textBbzid.Text.ToString() + "'";
            mc.ExeUpdate(ssql);
            

        }

        private void button10_Click(object sender, EventArgs e) //保存记录
        {

            this.InsertJLToMysql();
        }

        private void button7_Click(object sender, EventArgs e) //生成报账记录表,excel表
        {
            MysqlConnector mc = new MysqlConnector();
            mc.SetServer("127.0.0.1");
            mc.SetUserID("cwsh6");
            mc.SetPassword("1234");
            mc.SetDataBase("TQXZXXCWSHXT");

            //string ssql1 = "select bzxx,bzid,zflkm,ifnull(mxflkm,'合计：') as mxflkm,ifnull(sfxj,'小计') as sfxj ,";
            //string ssql2 = "ifnull(sfhg,'小计') as sfhg,sum(pjzs) as pjzs,sum(pjje) as pjje from bzjilu where bzid='" + this.textBbzid.Text.ToString() + "'" + " group by mxflkm,sfxj,sfhg with rollup"; //from bzjilu where bzid='"+this.textBbzid.Text.ToString()+"'"+groupBox1 
            //string ssql = ssql1 + ssql2;
            //MessageBox.Show(ssql);
            //this.dataGridView1.Rows.Clear();
            string ssql_pjzs_all = "select mxflkm,sum(pjzs) as pjzs from bzjilu where bzid='"+this.textBbzid.Text+"'"+" group by mxflkm";// 取出各个科目下的总张数
            MySqlDataReader hzjg = mc.ExeQuery(ssql_pjzs_all);
            if (!hzjg.Read())
            {
                MessageBox.Show("请保存数据后再生成报账申请表");
                return;
            }
            ExcelEditHelper do_excel = new ExcelEditHelper(); //生成操作excel的类
            
            do_excel.Open("c:\\MODE.xlsx");// 绝对路径
            
            do_excel.ws = do_excel.GetSheet("Sheet3");//获取表格方式
            //do_excel.SetCellValue(do_excel.ws, 1, 1, "tt"); 给单元格赋值方式

            //do_excel.wbs.Application.Visible = true; //设置此项可以让excel显示出来


            

                while (hzjg.Read())
                {
                    try
                    {
                        int i = 0;
                        //if(hzjg.GetString(0))
                        //do_excel.SetCellValue(do_excel.ws, 5, 3 + i, hzjg.GetString(0));
                        //do_excel.SetCellValue(do_excel.ws, 6, 3 + i, hzjg.GetDouble(1));
                        MessageBox.Show(hzjg.GetString(0));
                        MessageBox.Show(hzjg.GetString(1));
                        
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                        return;
                    }
                }
            
            do_excel.wbs.Application.Visible = true;
  
           


        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)//显示保存记录
        {
            this.dataGridView1.Rows.Clear(); //重新显示数据库中此次报账的记录，根据报账ID显示保存记录
            string ssql = "select * from bzjilu where bzid='" + this.textBbzid.Text.ToString() + "'";
            MysqlConnector mc = new MysqlConnector();
            mc.SetServer("127.0.0.1");
            mc.SetUserID("cwsh6");
            mc.SetPassword("1234");
            mc.SetDataBase("TQXZXXCWSHXT");

            MySqlDataReader hzjg = mc.ExeQuery(ssql);
            
            while (hzjg.Read())
            {
                int index = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[index].Cells[0].Value = hzjg.GetString(0);//报账学校
                this.dataGridView1.Rows[index].Cells[1].Value = hzjg.GetString(1);//报账ID
                this.dataGridView1.Rows[index].Cells[2].Value = hzjg.GetString(2);//总分类科目if
                this.dataGridView1.Rows[index].Cells[3].Value = hzjg.GetString(3);//明细分类科目
                //if (hzjg.GetString(3) == "合计：")
                //    this.dataGridView1.Rows[index].Cells[3].Style.ForeColor = Color.Red;
                this.dataGridView1.Rows[index].Cells[4].Value = hzjg.GetString(6);//是否现金
                //if (hzjg.GetString(4) == "小计")
                //    this.dataGridView1.Rows[index].Cells[4].Style.ForeColor = Color.Red;
                this.dataGridView1.Rows[index].Cells[5].Value = hzjg.GetString(7);//是否合格
                //if (hzjg.GetString(5) == "小计")
                //    this.dataGridView1.Rows[index].Cells[5].Style.ForeColor = Color.Red;

                this.dataGridView1.Rows[index].Cells[6].Value = hzjg.GetUInt32(4);//票据张数
                this.dataGridView1.Rows[index].Cells[7].Value = hzjg.GetDouble(5);//票据金额
                this.dataGridView1.Rows[index].Cells[8].Value = hzjg.GetString(8);//备注

            }
            

        }

        private void button11_Click(object sender, EventArgs e) //更新保存记录，把保存后修改的记录重新保存
        {
            MysqlConnector mc = new MysqlConnector();
            mc.SetServer("127.0.0.1");
            mc.SetUserID("cwsh6");
            mc.SetPassword("1234");
            mc.SetDataBase("TQXZXXCWSHXT");
            string ssql = "delete from bzjilu where bzid='" + this.textBbzid.Text.ToString() + "'";
            mc.ExeUpdate(ssql); //先把旧的数据从数据库中清除
            //然后再保存更新后的数据
            this.InsertJLToMysql();

        }
    }



}
