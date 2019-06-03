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
            this.comboBox1.SelectedIndex = 0;
        }

        private void listBox1_Click(object sender, EventArgs e)
        {
            this.selectedSchool = this.listBox1.SelectedItem.ToString().Substring(2);
            this.lb_welcome.Text = "欢迎" + this.selectedSchool + "前来报账！";
            //报账ID为年月日+日时分，除年之外，每一项都占2位，
            this.textBbzid.Text = DateTime.Now.ToString("yyyy-MM-dd hh:mm:mm");        // 2008-09-04 

            //在jfjlb中显示此学校的各项经费，及收入支出余额
            this.jfjlb.Rows.Clear(); //重新显示数据库中此次报账的记录，根据报账ID显示保存记录
            string ssql = "select * from xxje where bzxx='" + this.selectedSchool + "'";
            MysqlConnector mc = new MysqlConnector();
            mc.SetServer("192.168.78.189");
            mc.SetUserID("root");
            mc.SetPassword("root");
            mc.SetDataBase("TQXZXXCWSHXT");

            MySqlDataReader jejl = mc.ExeQuery(ssql); //jejl为从数据库表xxje中读取的记录
            try
            {

                while (jejl.Read())
                {
                    int index = this.jfjlb.Rows.Add();
                    this.jfjlb.Rows[index].Cells[0].Value = jejl.GetString(1);//报账学校
                    this.jfjlb.Rows[index].Cells[1].Value = jejl.GetString(2);//报账ID
                    this.jfjlb.Rows[index].Cells[2].Value = jejl.GetString(3);//总分类科目if
                    this.jfjlb.Rows[index].Cells[3].Value = jejl.GetString(4);//明细分类科目
                    

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            

            //清除GridView中的内容
            dataGridView1.Rows.Clear();


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

            textBzs.Focus();
            
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

                         this.dataGridView1.Rows[index].Cells[8].Value = this.comboBox1.SelectedItem.ToString();
                         this.dataGridView1.Rows[index].Cells[9].Value = this.textBbz.Text;//增加开支来源后，8是开支来源记录，9才是备注，20190603



                         this.textBzs.Text = "";
                         this.textBje.Text = "";
                     }
                 }


            textBzs.Focus();

                
            
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
        private bool InsertJLToMysql() //把datagrideview1中的记录数据插入mysql
        {
            int nRows=1;
            if (this.dataGridView1.Rows.Count > 1)
                nRows = this.dataGridView1.Rows.Count - 1;
            else
                return false;
            for (int i = 0; i < nRows; i++)
            {
                if (this.dataGridView1.Rows[i].Cells[5].Value==null)
                {
                    nRows = i;//为了防止用户点击了汇总后，又点击保存记录时把汇总项加上。
                    break;
                }
            }

            MysqlConnector mc = new MysqlConnector();
            mc.SetServer("192.168.78.189");
            mc.SetUserID("root");
            mc.SetPassword("root");
            mc.SetDataBase("TQXZXXCWSHXT");
            
            for(int i=0;i<nRows;i++)
            {
                //string ssql = "insert into bzjilu values(" + "'" + this.dataGridView1.Rows[i].Cells[0].Value + "'" + "," + "'" 
                //    + this.dataGridView1.Rows[i].Cells[1].Value + "'" + "," + "'" + this.dataGridView1.Rows[i].Cells[2].Value + 
                //    "'" + "," + "'" + this.dataGridView1.Rows[i].Cells[3].Value + "'" + "," + this.dataGridView1.Rows[i].Cells[6].Value +
                //    "," + this.dataGridView1.Rows[i].Cells[7].Value + "," + "'" + this.dataGridView1.Rows[i].Cells[4].Value + "'" + "," +
                //    "'" + this.dataGridView1.Rows[i].Cells[5].Value + "'" + "," + "'" + this.dataGridView1.Rows[i].Cells[8].Value + "'" + ")";
                //增加开支来源字段后，重新设置插入顺序
                string ssql = "insert into bzjilu values(" + "'" + this.dataGridView1.Rows[i].Cells[0].Value + "'" + "," + "'"
                    + this.dataGridView1.Rows[i].Cells[1].Value + "'" + "," + "'" + this.dataGridView1.Rows[i].Cells[2].Value +
                    "'" + "," + "'" + this.dataGridView1.Rows[i].Cells[3].Value + "'" + "," + this.dataGridView1.Rows[i].Cells[6].Value +
                    "," + this.dataGridView1.Rows[i].Cells[7].Value + "," + "'" + this.dataGridView1.Rows[i].Cells[4].Value + "'" + "," +
                    "'" + this.dataGridView1.Rows[i].Cells[5].Value + "'" + "," + "'" + this.dataGridView1.Rows[i].Cells[8].Value + "'" +//8是开支来源
                    "," + "'" + this.dataGridView1.Rows[i].Cells[9].Value + "'"  //9是备注                 
                    + ")";
                //MySqlDataAdapter reader = mc.ExeQuery(ssql);
                MessageBox.Show(ssql);

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
            mc.SetServer("192.168.78.189");
            mc.SetUserID("root");
            mc.SetPassword("root");
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





            this.dataGridView1.Rows.Add();//加一行空行，让原始记录和汇总记录分开
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

            //MysqlConnector mc = new MysqlConnector();
            //mc.SetServer("192.168.78.189");
            //mc.SetUserID("cwsh6");
            //mc.SetPassword("1234");
            //mc.SetDataBase("TQXZXXCWSHXT");
            //string ssql = "delete from bzjilu where bzid='" + this.textBbzid.Text.ToString() + "'";
            //mc.ExeUpdate(ssql);
            

        }

        private void button10_Click(object sender, EventArgs e) //保存记录
        {

            this.InsertJLToMysql();
        }

        private void button7_Click(object sender, EventArgs e) //生成报账记录表,excel表
        {
            string ssql_pjzs_all = "select ifnull(mxflkm,'合计：') as mxflkm,sum(pjzs) as pjzs， from bzjilu where bzid='" + 
                this.textBbzid.Text + "'" + " group by mxflkm with rollup";
            createBZJLB(ssql_pjzs_all,false, 6);//填写张数

            string ssql_pjje_all = "select ifnull(mxflkm,'合计：') as mxflkm,sum(pjje) as pjje， from bzjilu where bzid='" +
                this.textBbzid.Text + "'" + " and sfxj='现金' group by mxflkm with rollup";
            createBZJLB(ssql_pjje_all, false, 7);//填写票据现金金额



           

        }
        private void createBZJLB(string msql,bool sh,int nrow) //调用此函数生成报账记录表,sh表示是申请的金额还是审核的金额，false表示申请金额，true表示
                                                            //审核金额，nrow表示在第几行插入相应数字,在申请栏中6是张数，7是现金，8是转账，9是总额。
        {                                              //在审核后栏中，10是审核后张数，11是审核后现金，12是审核后转账，13审核后总金额。
            MysqlConnector mc = new MysqlConnector();
            mc.SetServer("192.168.78.189");
            mc.SetUserID("root");
            mc.SetPassword("root");
            mc.SetDataBase("TQXZXXCWSHXT");

            string ssql = msql; // "select ifnull(mxflkm,'合计：') as mxflkm,sum(pjje) as pjje， from bzjilu where bzid='" + this.textBbzid.Text + "'" + " group by mxflkm with rollup";// 取出各个科目下的总张数
            //string ssql_xj_all = "select ifnull(mxflkm,'合计：') as mxflkm,sum(pjje) as pjje， from bzjilu where bzid='" + this.textBbzid.Text + "'" + " group by mxflkm with rollup";
            MySqlDataReader hzjg = mc.ExeQuery(ssql); //汇总票据张数结果
            //MySqlDataReader xj = mc.ExeQuery(ssql_xj_all);//汇总现金结果

            if (!hzjg.Read())
            {
                MessageBox.Show("请保存数据后再生成报账申请表");
                return;
            }
            ExcelEditHelper do_excel = new ExcelEditHelper(); //生成操作excel的类

            do_excel.Open("D:\\MODE.xlsx");// 绝对路径

            do_excel.ws = do_excel.GetSheet("Sheet3");//获取表格方式

            int i = 0;
            do  //循环输出各个科目的汇总张数及金额
            {
                try
                {

                    if (hzjg.GetString(0) != "合计：") //表示到了汇总栏了，合计栏不用填
                    {
                        if (sh == false)
                        {
                            do_excel.SetCellValue(do_excel.ws, 5, 3 + i, hzjg.GetString(0));
                            do_excel.SetCellValue(do_excel.ws, nrow, 3 + i, hzjg.GetDouble(1));
                            //do_excel.SetCellValue(do_excel.ws, 7, 3 + i, xj.GetDouble(1));
                        }
                        else
                        {
                            do_excel.SetCellValue(do_excel.ws, 10, 3 + i, hzjg.GetString(0));
                            do_excel.SetCellValue(do_excel.ws, nrow, 3 + i, hzjg.GetDouble(1));

                        }
                    }
                   
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    return;
                }
                //MessageBox.Show(hzjg.GetString(0)); 20181106测试，找出问题，每执行一次read()，就会换一行。
                //MessageBox.Show(hzjg.GetString(1)); 2018-11-06 03:22:22
                i++;
            } while (hzjg.Read());



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
            mc.SetServer("192.168.78.189");
            mc.SetUserID("root");
            mc.SetPassword("root");
            mc.SetDataBase("TQXZXXCWSHXT");

            MySqlDataReader hzjg = mc.ExeQuery(ssql);
            try
            {

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
                    this.dataGridView1.Rows[index].Cells[8].Value = hzjg.GetString(8);//新加开支来源,20190603
                    this.dataGridView1.Rows[index].Cells[9].Value = hzjg.GetString(9);//新变为9，才是备注

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            

        }

        private void button11_Click(object sender, EventArgs e) //更新保存记录，把保存后修改的记录重新保存
        {
            MysqlConnector mc = new MysqlConnector();
            mc.SetServer("192.168.78.189");
            mc.SetUserID("root");
            mc.SetPassword("root");
            mc.SetDataBase("TQXZXXCWSHXT");
            string ssql = "delete from bzjilu where bzid='" + this.textBbzid.Text.ToString() + "'";
            mc.ExeUpdate(ssql); //先把旧的数据从数据库中清除
            //然后再保存更新后的数据
            this.InsertJLToMysql();

        }

        private void button12_Click(object sender, EventArgs e)
        {
            try
                {
            this.dataGridView1.Rows.Remove(this.dataGridView1.CurrentRow);
                }
                catch{
                }
        }

        private void textBzs_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                SendKeys.Send("{Tab}");
            } 

        }

        private void textBje_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                SendKeys.Send("{Tab}");
            } 

        }

        private void checkBsfxj_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                SendKeys.Send("{Tab}");
            }

        }

        private bool saveJFJE() //保存学校经费记录项目
        {
            int nRows = 1;
            if (this.jfjlb.Rows.Count > 1)
                nRows = this.jfjlb.Rows.Count - 1;
            else
                return false;
           

            MysqlConnector mc = new MysqlConnector();
            mc.SetServer("192.168.78.189");
            mc.SetUserID("root");
            mc.SetPassword("root");
            mc.SetDataBase("TQXZXXCWSHXT");
            //删除表中该学校原有记录，重新生成该学校新的记录
            string sqld = "delete from xxje where bzxx='" + this.selectedSchool + "'";
            mc.ExeUpdate(sqld);

            for (int i = 0; i < nRows; i++)
            {
                string ssql = "insert into xxje values(" + "'" + this.selectedSchool + "'" + "," + "'" + this.jfjlb.Rows[i].Cells[0].Value + "'" + "," + "'"
                    + this.jfjlb.Rows[i].Cells[1].Value + "'" + "," + "'" + this.jfjlb.Rows[i].Cells[2].Value +
                    "'" + "," + "'" + this.jfjlb.Rows[i].Cells[3].Value + "'"+ ")";
                //MySqlDataAdapter reader = mc.ExeQuery(ssql);

               // MessageBox.Show(ssql); 显示插入到数据库的sql句子

                //mc.ExeQuery(ssql);
                mc.ExeUpdate(ssql);

            }

            MessageBox.Show("保存完毕");
           

            return true;
        }
        private void button6_Click(object sender, EventArgs e)  //保存经费修改，经费表中添加、修改或删除后，保存到表中
        {//在数据库中新建了一个xxje表，记录学校金额，表中共有五个字段，分别为
          //bzxx varchar(30),jfly varchar(50),zsr double,zzc double,ye  。

            saveJFJE();

        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                this.jfjlb.Rows.Remove(this.jfjlb.CurrentRow);
            }
            catch
            {
            }
        }
    }



}
