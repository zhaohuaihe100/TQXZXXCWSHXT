using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TQXZXXCWSHXT
{
  public   struct ZsJe
    {
        public string leixing;//记录张数金额的类型，共7种
        public double zs; //记录张数
        public double je; //记录金额
    };
    public class RecordHZ  //定义一个处理汇总数据的类
    {
        private string schoolName;//报账学校
        private string bzID;  //报账ID
        private string zflK;  //报账总分类科目
        private string mxflK;  //报账明细分类科目
        //private string 

      private ZsJe[] zsje7; //张数和金额共有7种类型要记录，总金额0、总现金1、总转账2、合格现金3、
        //合格转账4、不合格现金5、不合格转账6，。

        public RecordHZ() //构造函数初始化变量
        {
            this.schoolName = "";
            this.bzID="";
            this.zflK = "";
            this.mxflK = "";
            zsje7 = new ZsJe[7];//定义7种类型的张数金额记录

            for (int i = 0; i < 7; i++)
            {
                switch(i)
                {
                    case 0:
                        zsje7[i].leixing = "总金额";
                        break;
                    case 1:
                        zsje7[i].leixing = "总现金金额";
                        break;
                    case 2:
                        zsje7[i].leixing = "总转账金额";
                        break;
                    case 3:
                        zsje7[i].leixing = "合格现金金额";
                        break;
                    case 4:
                        zsje7[i].leixing = "合格转账金额";
                        break;
                    case 5:
                        zsje7[i].leixing = "不合格现金金额";
                        break;
                    case 6:
                        zsje7[i].leixing = "不合格转账金额";
                        break;
                }
                

                zsje7[i].zs = 0.0;
                zsje7[i].je = 0.0;//初始化张数和金额

            }


            
        }
        public void InitRecordHZ() //初始化记录中的各个参数，当分类汇总
        {
            this.zflK = "";
            this.mxflK = "";

            for (int i = 0; i < 7; i++)
            {
                zsje7[i].leixing = "";
                zsje7[i].zs = 0.0;
                zsje7[i].je = 0.0;//初始化张数和金额

            }

        }

        public void AddZsJe(string stype, int tzs, double tje)
        {
              switch (stype)
                {
                    case "总金额":
                       // zsje7[0].leixing = "总金额";
                        zsje7[0].zs = zsje7[0].zs + tzs;
                        zsje7[0].je = zsje7[0].je + tje;
                        break;
                    case "总现金金额":
                       // zsje7[1].leixing = "总现金金额";
                       zsje7[1].zs = zsje7[1].zs + tzs;
                        zsje7[1].je = zsje7[1].je + tje;
                        break;
                    case "总转账金额":
                       // zsje7[2].leixing = "总转账金额";
                       zsje7[2].zs = zsje7[2].zs + tzs;
                        zsje7[2].je = zsje7[2].je + tje;
                        break;
                    case "合格现金金额":
                        //zsje7[3].leixing = "合格现金金额";
                        zsje7[3].zs = zsje7[3].zs + tzs;
                        zsje7[3].je = zsje7[3].je + tje;
                        break;
                    case "合格转账金额":
                        //zsje7[4].leixing = "合格转账金额";
                       zsje7[4].zs = zsje7[4].zs + tzs;
                        zsje7[4].je = zsje7[4].je + tje;
                        break;
                    case "不合格现金金额":
                        //zsje7[5].leixing = "不合格现金金额";
                       zsje7[5].zs = zsje7[5].zs + tzs;
                        zsje7[5].je = zsje7[5].je + tje;
                        break;
                    case "不合格转账金额":
                        //zsje7[6].leixing = "不合格转账金额";
                        zsje7[6].zs = zsje7[6].zs + tzs;
                        zsje7[6].je = zsje7[6].je + tje;
                        break;
                }
            

        }
        


    }

    //public class Connect
}
