using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Windows.Forms;
//using System.Data.SqlClient;
using System.Configuration;
using MySql.Data;
using MySql.Data.MySqlClient;
//using Microsoft.Office.Core;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
namespace TQXZXXCWSHXT
{
    class ExcelEditHelper
    {
        public string mFilename;
        public Excel.Application app;
        public Workbooks wbs;
        public Workbook wb;
        public Worksheets wss;//工作表集合
        public Worksheet ws;//工作表

        /// <summary>
        /// 创建一个Excel对象
        /// </summary>
        public void Create()
        {
            app = new Excel.Application();
            wbs = app.Workbooks;
            wb = wbs.Add(true);
        }

        /// <summary>
        /// 打开一个Excel文件
        /// </summary>
        /// <param name="FileName">excel文件名，包括文件路径</param>
        public void Open(string FileName)
        {
            try
            {
                object missing = System.Reflection.Missing.Value;
                app = new Excel.Application();
                wbs = app.Workbooks;

                wb = wbs.Add(FileName);
                wb = wbs.Open(FileName, missing, true, missing, missing, missing, missing, missing, missing, true, missing, missing, missing, missing, missing);
                mFilename = FileName;
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return;
            }
        }

        /// <summary>
        /// 获取一个工作表
        /// </summary>
        /// <param name="SheetName">工作表名称</param>
        /// <returns></returns>
        public Worksheet GetSheet(string SheetName)
        {
            Worksheet s = (Worksheet)wb.Worksheets[SheetName];
            return s;
        }

        /// <summary>
        /// 添加一个工作表
        /// </summary>
        /// <param name="SheetName">工作表名称</param>
        /// <returns></returns>
        public Worksheet AddSheet(string SheetName)
        {
            Worksheet s = (Worksheet)wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            s.Name = SheetName;
            return s;
        }

        /// <summary>
        /// 复制并添加一个工作表
        /// </summary>
        /// <param name="OldSheetName"> 被复制工作表</param>
        /// <param name="NewSheetName">新表</param>
        public void CloneSheet(string OldSheetName, string NewSheetName)
        {
            Microsoft.Office.Interop.Excel.Worksheet oldSheet = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[OldSheetName];
            oldSheet.Copy(oldSheet, Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet s = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[OldSheetName + " (2)"];
            s.Name = NewSheetName;
        }

        /// <summary>
        /// 删除一个工作表
        /// </summary>
        /// <param name="SheetName">工作表名称</param>
        public void DelSheet(string SheetName)
        {
            ((Worksheet)wb.Worksheets[SheetName]).Delete();
        }

        /// <summary>
        /// 重命名一个工作表
        /// </summary>
        /// <param name="OldSheetName">被替换名称</param>
        /// <param name="NewSheetName">替换名称</param>
        /// <returns></returns>
        public Worksheet ReNameSheet(string OldSheetName, string NewSheetName)
        {
            Worksheet s = (Worksheet)wb.Worksheets[OldSheetName];
            s.Name = NewSheetName;
            return s;
        }

        /// <summary>
        /// 重命名一个工作表
        /// </summary>
        /// <param name="Sheet">被替换工作表</param>
        /// <param name="NewSheetName">替换名称</param>
        /// <returns></returns>
        public Worksheet ReNameSheet(Worksheet Sheet, string NewSheetName)
        {
            Sheet.Name = NewSheetName;
            return Sheet;
        }

        /// <summary>
        /// 设置单元格的值
        /// </summary>
        /// <param name="ws">工作表</param>
        /// <param name="x">行标</param>
        /// <param name="y">列标</param>
        /// <param name="value">数据</param>
        public void SetCellValue(Worksheet ws, int x, int y, object value)
        {
            ws.Cells[x, y] = value;
        }

        /// <summary>
        /// 设置单元格的值
        /// </summary>
        /// <param name="ws">工作表名称</param>
        /// <param name="x">行标</param>
        /// <param name="y">列标</param>
        /// <param name="value">数据</param>
        public void SetCellValue(string ws, int x, int y, object value)
        {
            GetSheet(ws).Cells[x, y] = value;
        }

        /// <summary>
        /// 设置单元格属性
        /// </summary>
        /// <param name="ws">工作表</param>
        /// <param name="Startx">起始行标</param>
        /// <param name="Starty">起始列标</param>
        /// <param name="Endx">终止行标</param>
        /// <param name="Endy">终止列标</param>
        /// <param name="size">字体大小</param>
        /// <param name="name">字体</param>
        /// <param name="color">颜色</param>
        /// <param name="HorizontalAlignment">对齐方式</param>
        public void SetCellProperty(Worksheet ws, int Startx, int Starty, int Endx, int Endy, int size, string name, Constants color, Constants HorizontalAlignment)
        {
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Name = name;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Size = size;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Color = color;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).HorizontalAlignment = HorizontalAlignment;
        }

        /// <summary>
        /// 设置单元格属性
        /// </summary>
        /// <param name="ws">工作表名称</param>
        /// <param name="Startx">起始行标</param>
        /// <param name="Starty">起始列标</param>
        /// <param name="Endx">终止行标</param>
        /// <param name="Endy">终止列标</param>
        /// <param name="size">字体大小</param>
        /// <param name="name">字体</param>
        /// <param name="color">颜色</param>
        /// <param name="HorizontalAlignment">对齐方式</param>
        public void SetCellProperty(string wsn, int Startx, int Starty, int Endx, int Endy, int size, string name, Constants color, Constants HorizontalAlignment)
        {
            Worksheet ws = GetSheet(wsn);
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Name = name;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Size = size;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Color = color;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).HorizontalAlignment = HorizontalAlignment;
        }

        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="ws">工作表</param>
        /// <param name="x1">起始行标</param>
        /// <param name="y1">起始列标</param>
        /// <param name="x2">终止行标</param>
        /// <param name="y2">终止列标</param>
        public void UniteCells(Worksheet ws, int x1, int y1, int x2, int y2)
        {
            ws.get_Range(ws.Cells[x1, y1], ws.Cells[x2, y2]).Merge(Type.Missing);
        }

        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="ws">工作表名称</param>
        /// <param name="x1">起始行标</param>
        /// <param name="y1">起始列标</param>
        /// <param name="x2">终止行标</param>
        /// <param name="y2">终止列标</param>
        public void UniteCells(string ws, int x1, int y1, int x2, int y2)
        {
            GetSheet(ws).get_Range(GetSheet(ws).Cells[x1, y1], GetSheet(ws).Cells[x2, y2]).Merge(Type.Missing);
        }

        /// <summary>
        /// 将内存中数据表格插入到Excel指定工作表的指定位置
        /// </summary>
        /// <param name="dt">数据表</param>
        /// <param name="ws">工作表名称</param>
        /// <param name="startX">起始行标</param>
        /// <param name="startY">起始列标</param>
        public void InsertTable(System.Data.DataTable dt, string ws, int startX, int startY)
        {
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= dt.Columns.Count - 1; j++)
                {
                    GetSheet(ws).Cells[startX + i, j + startY] = dt.Rows[i][j];
                }
            }
        }

        /// <summary>
        /// 将内存中数据表格插入到Excel指定工作表的指定位置
        /// </summary>
        /// <param name="dt">数据表</param>
        /// <param name="ws">工作表名称</param>
        /// <param name="startX">起始行标</param>
        /// <param name="startY">起始列标</param>
        public void InsertTable(System.Data.DataTable dt, Worksheet ws, int startX, int startY)
        {
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= dt.Columns.Count - 1; j++)
                {
                    ws.Cells[startX + i, j + startY] = dt.Rows[i][j];
                }
            }
        }

        /// <summary>
        /// 保存文档
        /// </summary>
        /// <returns></returns>
        public bool Save()
        {
            if (mFilename == "")
            {
                return false;
            }
            else
            {
                try
                {
                    wb.Save();
                    return true;
                }
                catch (Exception ex)
                {
                    return false;
                }
            }
        }

        /// <summary>
        /// 文档另存为
        /// </summary>
        /// <param name="FileName">文件名（包含路径）</param>
        /// <returns></returns>
        public bool SaveAs(object FileName)
        {
            try
            {
                wb.SaveAs(FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        /// <summary>
        /// 关闭一个Excel对象，销毁对象
        /// </summary>
        public void Close()
        {
            wb.Close(Type.Missing, Type.Missing, Type.Missing);
            wbs.Close();
            app.Quit();
            wb = null;
            wbs = null;
            app = null;
            GC.Collect();
        }
        public void Close_wb()
        {
            Excel.Workbook workbook = this.app.Workbooks.get_Item(1);  //this.Application.Workbooks.get_Item(fileName);
            workbook.Close(false, Type.Missing, Type.Missing);
        }
    }
}
