using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;


namespace FCfwz
{
    public partial class Form1 : Form
    {
        public SQLServer MySQLServer = new SQLServer();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            string ret = GetAppSettings();
            if (ret == "")
            {
                this.textBox1.Text = MySQLServer.SQL_Name;
                this.textBox2.Text = MySQLServer.SQL_ID;
                this.textBox3.Text = MySQLServer.SQL_PassWord;
                this.textBox4.Text = MySQLServer.SQL_DataBase;
            }
            else
            {
                MySQLServer.SQL_Name = "";
                MySQLServer.SQL_ID = "";
                MySQLServer.SQL_PassWord = "";
                MySQLServer.SQL_DataBase = "";
                this.textBox1.Text = "";
                this.textBox2.Text = "";
                this.textBox3.Text = "";
                this.textBox4.Text = "";
            }
            MySQLServer.SelfConn = false;
            MySQLServer.TestConnection();
            this.comboBox1.Items.Clear();
            List<ComboBoxItem> Departments = GetAllDepartment();
            foreach (ComboBoxItem Department in Departments)
            {
                this.comboBox1.Items.Add(Department);
            }
            if (this.comboBox1.Items.Count > 0) { this.comboBox1.SelectedIndex = 0; }

            this.dateTimePicker1.Value = DateTime.Now.AddDays(-7);
            this.dateTimePicker2.Value = DateTime.Now; 

            this.comboBox2.Items.Add("库房");
            this.comboBox2.Items.Add("类别");
            this.comboBox2.SelectedIndex = 0;

            this.comboBox3.Items.Add("处置科室");
            this.comboBox3.Items.Add("开方科室");
            this.comboBox3.SelectedIndex = 0;
            this.toolStripStatusLabel1.Text = "";
            Application.DoEvents();
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            this.splitContainer1.SplitterDistance = 320;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (this.splitContainer1.SplitterDistance == 320)
            {
                this.splitContainer1.SplitterDistance = 50;
                this.button1.Text = ">>";
                this.groupBox1.Visible = false;
                this.groupBox2.Visible = false;
                this.groupBox3.Visible = false;
                this.groupBox4.Visible = false;
                this.button4.Visible = false;
                this.button5.Visible = false;
                this.button6.Visible = false;

            }
            else
            {
                this.splitContainer1.SplitterDistance = 320;
                this.button1.Text = "配置项设置，单击可以隐藏";
                this.groupBox1.Visible = true;
                this.groupBox2.Visible = true;
                this.groupBox3.Visible = true;
                this.groupBox4.Visible = true;
                this.button4.Visible = true;
                this.button5.Visible = true;
                this.button6.Visible = true;
            }
        }

        private void toolTip1_Popup(object sender, PopupEventArgs e)
        {

        }

        private void button1_MouseEnter(object sender, EventArgs e)
        {
            if (this.splitContainer1.SplitterDistance == 50)
            {
                this.toolTip1.IsBalloon = false;
                this.toolTip1.UseFading = true;
                this.toolTip1.Show("配置项设置，单击可以隐藏", this.button1);
            }
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            this.toolTip1.Hide(this.button1);     //隐藏提示窗口
        }
        //测试
        private void button3_Click(object sender, EventArgs e)
        {
            MySQLServer.SQL_Name = this.textBox1.Text;
            MySQLServer.SQL_ID = this.textBox2.Text;
            MySQLServer.SQL_PassWord = this.textBox3.Text;
            MySQLServer.SQL_DataBase = this.textBox4.Text;
            MySQLServer.SelfConn = false;
            MySQLServer.TestConnection();
            if (MySQLServer.SelfConn == true)
            {
                MessageBox.Show("远程数据库连接成功！！！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            else
            {
                MessageBox.Show("远程数据库连接失败！！！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //保存
        private void button2_Click(object sender, EventArgs e)
        {
            MySQLServer.SQL_Name = this.textBox1.Text;
            MySQLServer.SQL_ID = this.textBox2.Text;
            MySQLServer.SQL_PassWord = this.textBox3.Text;
            MySQLServer.SQL_DataBase = this.textBox4.Text;
            if (MySQLServer.SQL_Name == "" || MySQLServer.SQL_ID == "" || MySQLServer.SQL_DataBase == "")
            {
                MessageBox.Show("请填写准确的参数，测试通过后再保存！！！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                string ret = SetAppSettings();
                if (ret != "")
                {
                    MessageBox.Show(ret, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show("配置信息保存成功！！！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
        }
        //查询
        private void button4_Click(object sender, EventArgs e)
        {
            if (MySQLServer.SelfConn == true)
            {
                List<s_m_total> m = new List<s_m_total>();
                List<s_n_total> n = new List<s_n_total>();

                int t = (this.checkBox1.Checked == true ? 1 : 0);
                DataTable dt = new DataTable();

                if (this.radioButton1.Checked == true)
                {
                    if (this.comboBox2.SelectedIndex == 0)
                    {
                        //根据药库分
                        m = Get_m1_total();
                        #region 仓库
                        dt.Columns.Add("仓库", System.Type.GetType("System.String"));
                        dt.Columns.Add("仓库数量", System.Type.GetType("System.Int32"));
                        dt.Columns.Add("仓库金额", System.Type.GetType("System.Decimal"));
                        dt.Columns.Add("药品", System.Type.GetType("System.String"));
                        dt.Columns.Add("规格", System.Type.GetType("System.String"));
                        dt.Columns.Add("单位", System.Type.GetType("System.String"));
                        dt.Columns.Add("单价", System.Type.GetType("System.Decimal"));
                        dt.Columns.Add("药品数量", System.Type.GetType("System.Int32"));
                        dt.Columns.Add("药品金额", System.Type.GetType("System.Decimal"));
                        if (t == 0)
                        {
                            dt.Columns.Add("部门", System.Type.GetType("System.String"));
                            dt.Columns.Add("部门数量", System.Type.GetType("System.Int32"));
                            dt.Columns.Add("部门金额", System.Type.GetType("System.Decimal"));
                            dt.Columns.Add("医师", System.Type.GetType("System.String"));
                            dt.Columns.Add("医师数量", System.Type.GetType("System.Int32"));
                            dt.Columns.Add("医师金额", System.Type.GetType("System.Decimal"));
                        }
                        foreach (s_m_total item in m)
                        {
                            if (t == 0)
                            {
                                dt.Rows.Add(item.库名, Convert.ToInt32(item.库码数量), item.库码金额,
                                        item.药品名称, item.规格, item.单位, item.单价, Convert.ToInt32(item.药品数量), item.药品金额,
                                        item.部门名称, Convert.ToInt32(item.部门数量), item.部门金额,
                                        item.医师名称, Convert.ToInt32(item.医师数量), item.医师金额);
                            }
                            else
                            {
                                dt.Rows.Add(item.库名, Convert.ToInt32(item.库码数量), item.库码金额,
                                        item.药品名称, item.规格, item.单位, item.单价, Convert.ToInt32(item.药品数量), item.药品金额);
                            }
                        }
                        #endregion
                    }
                    else
                    {
                        //根据药类分
                        m = Get_m2_total();
                        #region 类别
                        dt.Columns.Add("类别", System.Type.GetType("System.String"));
                        dt.Columns.Add("类别数量", System.Type.GetType("System.Int32"));
                        dt.Columns.Add("类别金额", System.Type.GetType("System.Decimal"));
                        dt.Columns.Add("药品", System.Type.GetType("System.String"));
                        dt.Columns.Add("规格", System.Type.GetType("System.String"));
                        dt.Columns.Add("单位", System.Type.GetType("System.String"));
                        dt.Columns.Add("单价", System.Type.GetType("System.Decimal"));
                        dt.Columns.Add("药品数量", System.Type.GetType("System.Int32"));
                        dt.Columns.Add("药品金额", System.Type.GetType("System.Decimal"));
                        if (t == 0)
                        {
                            dt.Columns.Add("部门", System.Type.GetType("System.String"));
                            dt.Columns.Add("部门数量", System.Type.GetType("System.Int32"));
                            dt.Columns.Add("部门金额", System.Type.GetType("System.Decimal"));
                            dt.Columns.Add("医师", System.Type.GetType("System.String"));
                            dt.Columns.Add("医师数量", System.Type.GetType("System.Int32"));
                            dt.Columns.Add("医师金额", System.Type.GetType("System.Decimal"));
                        }
                        foreach (s_m_total item in m)
                        {
                            if (t == 0)
                            {
                                dt.Rows.Add(item.类名, Convert.ToInt32(item.类码数量), item.类码金额,
                                        item.药品名称, item.规格, item.单位, item.单价, Convert.ToInt32(item.药品数量), item.药品金额,
                                        item.部门名称, Convert.ToInt32(item.部门数量), item.部门金额,
                                        item.医师名称, Convert.ToInt32(item.医师数量), item.医师金额);
                            }
                            else
                            {
                                dt.Rows.Add(item.类名, Convert.ToInt32(item.类码数量), item.类码金额,
                                        item.药品名称, item.规格, item.单位, item.单价, Convert.ToInt32(item.药品数量), item.药品金额);
                            }
                        }
                        #endregion
                    }
                    this.dataGridView1.DataSource = null;
                    this.dataGridView1.DataSource = dt;
                    MessageBox.Show("数据获取完毕！！！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                if (this.radioButton2.Checked == true)
                {
                    if (this.radioButton3.Checked == true)
                    {
                        if (this.comboBox3.SelectedIndex == 0)
                        {
                            //处置科室
                            n = Get_n1_total();
                            #region 科室
                            dt.Columns.Add("科室", System.Type.GetType("System.String"));
                            dt.Columns.Add("科室数量", System.Type.GetType("System.Int32"));
                            dt.Columns.Add("科室金额", System.Type.GetType("System.Decimal"));
                            dt.Columns.Add("项目", System.Type.GetType("System.String"));
                            dt.Columns.Add("项目数量", System.Type.GetType("System.Int32"));
                            dt.Columns.Add("项目金额", System.Type.GetType("System.Decimal"));
                            if (t == 0)
                            {
                                dt.Columns.Add("部门", System.Type.GetType("System.String"));
                                dt.Columns.Add("部门数量", System.Type.GetType("System.Int32"));
                                dt.Columns.Add("部门金额", System.Type.GetType("System.Decimal"));
                                dt.Columns.Add("医师", System.Type.GetType("System.String"));
                                dt.Columns.Add("医师数量", System.Type.GetType("System.Int32"));
                                dt.Columns.Add("医师金额", System.Type.GetType("System.Decimal"));
                            }
                            foreach (s_n_total item in n)
                            {
                                if (t == 0)
                                {
                                    dt.Rows.Add(item.科名, Convert.ToInt32(item.科码数量), item.科码金额,
                                            item.项目名称, Convert.ToInt32(item.项目数量), item.项目金额,
                                            item.部门名称, Convert.ToInt32(item.部门数量), item.部门金额,
                                            item.医师名称, Convert.ToInt32(item.医师数量), item.医师金额);
                                }
                                else
                                {
                                    dt.Rows.Add(item.科名, Convert.ToInt32(item.科码数量), item.科码金额,
                                            item.项目名称, Convert.ToInt32(item.项目数量), item.项目金额);
                                }
                            }
                            #endregion
                        }
                        else
                        {
                            //处方来源
                            n = Get_n2_total();
                            #region 部门
                            dt.Columns.Add("科室", System.Type.GetType("System.String"));
                            dt.Columns.Add("科室数量", System.Type.GetType("System.Int32"));
                            dt.Columns.Add("科室金额", System.Type.GetType("System.Decimal"));
                            dt.Columns.Add("项目", System.Type.GetType("System.String"));
                            dt.Columns.Add("项目数量", System.Type.GetType("System.Int32"));
                            dt.Columns.Add("项目金额", System.Type.GetType("System.Decimal"));
                            if (t == 0)
                            {
                                dt.Columns.Add("部门", System.Type.GetType("System.String"));
                                dt.Columns.Add("部门数量", System.Type.GetType("System.Int32"));
                                dt.Columns.Add("部门金额", System.Type.GetType("System.Decimal"));
                                dt.Columns.Add("医师", System.Type.GetType("System.String"));
                                dt.Columns.Add("医师数量", System.Type.GetType("System.Int32"));
                                dt.Columns.Add("医师金额", System.Type.GetType("System.Decimal"));
                            }
                            foreach (s_n_total item in n)
                            {
                                if (t == 0)
                                {
                                    dt.Rows.Add(item.科名, Convert.ToInt32(item.科码数量), item.科码金额,
                                            item.项目名称, Convert.ToInt32(item.项目数量), item.项目金额,
                                            item.部门名称, Convert.ToInt32(item.部门数量), item.部门金额,
                                            item.医师名称, Convert.ToInt32(item.医师数量), item.医师金额);
                                }
                                else
                                {
                                    dt.Rows.Add(item.科名, Convert.ToInt32(item.科码数量), item.科码金额,
                                            item.项目名称, Convert.ToInt32(item.项目数量), item.项目金额);
                                }
                            }
                            #endregion
                        }
                        this.dataGridView1.DataSource = null;
                        this.dataGridView1.DataSource = dt;
                        MessageBox.Show("数据获取完毕！！！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        //科室项目交叉
                        if (this.radioButton4.Checked == true)
                        {
                            Get_nx_total(1);
                        }
                        //项目科室交叉
                        if (this.radioButton5.Checked == true)
                        {
                            Get_nx_total(2);
                        }
                        //医师项目交叉
                        if (this.radioButton6.Checked == true)
                        {
                            Get_nx_total(3);
                        }
                        //项目医师交叉
                        if (this.radioButton7.Checked == true)
                        {
                            Get_nx_total(4);
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("远程数据库连接失败！！！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

            this.toolStripStatusLabel1.Text = "";
            Application.DoEvents();

        }
        //导出
        private void button5_Click(object sender, EventArgs e)
        {
            if (MySQLServer.SelfConn == true)
            {
                List<s_m_total> m = new List<s_m_total>();
                List<s_n_total> n = new List<s_n_total>();

                if (this.radioButton1.Checked == true)
                {

                    if (this.comboBox2.SelectedIndex == 0)
                    {
                        //根据药库分
                        m = Get_m1_total();
                        Excel_m1_total(m);
                    }
                    else
                    {
                        //根据药类分
                        m = Get_m2_total();
                        Excel_m2_total(m);
                    }
                }

                if (this.radioButton2.Checked == true)
                {
                    if (this.radioButton3.Checked == true)
                    {
                        if (this.comboBox3.SelectedIndex == 0)
                        {
                            //处置科室
                            n = Get_n1_total();
                            Excel_n1_total(n);
                        }
                        else
                        {
                            //处方来源
                            n = Get_n2_total();
                            Excel_n2_total(n);
                        }
                    }
                    else
                    {

                        //科室项目交叉
                        if (this.radioButton4.Checked == true)
                        {
                            Excel_nx_total(1);
                        }
                        //项目科室交叉
                        if (this.radioButton5.Checked == true)
                        {
                            Excel_nx_total(2);
                        }
                        //医师项目交叉
                        if (this.radioButton6.Checked == true)
                        {
                            Excel_nx_total(3);
                        }
                        //项目医师交叉
                        if (this.radioButton7.Checked == true)
                        {
                            Excel_nx_total(4);
                        }

                    }

                }
            }
            else
            {
                MessageBox.Show("远程数据库连接失败！！！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            this.toolStripStatusLabel1.Text = "";
            Application.DoEvents();

        }



        #region 2018-11-16
        public void Excel_m1_total(List<s_m_total> m)
        {

            int t = (this.checkBox1.Checked == true ? 1 : 0);
            DataTable dt = new DataTable();
            dt.Columns.Add("仓库", System.Type.GetType("System.String"));
            dt.Columns.Add("仓库数量", System.Type.GetType("System.Int32"));
            dt.Columns.Add("仓库金额", System.Type.GetType("System.Decimal"));
            dt.Columns.Add("药品", System.Type.GetType("System.String"));
            dt.Columns.Add("规格", System.Type.GetType("System.String"));
            dt.Columns.Add("单位", System.Type.GetType("System.String"));
            dt.Columns.Add("单价", System.Type.GetType("System.Decimal"));
            dt.Columns.Add("药品数量", System.Type.GetType("System.Int32"));
            dt.Columns.Add("药品金额", System.Type.GetType("System.Decimal"));
            if (t == 0)
            {
                dt.Columns.Add("部门", System.Type.GetType("System.String"));
                dt.Columns.Add("部门数量", System.Type.GetType("System.Int32"));
                dt.Columns.Add("部门金额", System.Type.GetType("System.Decimal"));
                dt.Columns.Add("医师", System.Type.GetType("System.String"));
                dt.Columns.Add("医师数量", System.Type.GetType("System.Int32"));
                dt.Columns.Add("医师金额", System.Type.GetType("System.Decimal"));
            }
            foreach (s_m_total item in m)
            {
                if (t == 0)
                {
                    dt.Rows.Add(item.库名, Convert.ToInt32(item.库码数量), item.库码金额,
                            item.药品名称, item.规格, item.单位, item.单价, Convert.ToInt32(item.药品数量), item.药品金额,
                            item.部门名称, Convert.ToInt32(item.部门数量), item.部门金额,
                            item.医师名称, Convert.ToInt32(item.医师数量), item.医师金额);
                }
                else
                {
                    dt.Rows.Add(item.库名, Convert.ToInt32(item.库码数量), item.库码金额,
                            item.药品名称, item.规格, item.单位, item.单价, Convert.ToInt32(item.药品数量), item.药品金额);
                }
            }
            IWorkbook workBook = new HSSFWorkbook();
            workBook = ExcelHelper.ToExcel(dt, "肥城市妇幼保健院药品销售(库房)分类明细");

            ISheet sheet1 = workBook.GetSheetAt(0);
            sheet1.SetColumnWidth(0, 20 * 256);     //仓库
            sheet1.SetColumnWidth(1, 12 * 256);     //数量
            sheet1.SetColumnWidth(2, 12 * 256);     //金额

            sheet1.SetColumnWidth(3, 30 * 256);     //药品名称
            sheet1.SetColumnWidth(4, 15 * 256);     //规格
            sheet1.SetColumnWidth(5, 6 * 256);      //单位
            sheet1.SetColumnWidth(6, 12 * 256);     //单价
            sheet1.SetColumnWidth(7, 12 * 256);     //数量
            sheet1.SetColumnWidth(8, 12 * 256);     //金额
            if (t == 0)
            {
                sheet1.SetColumnWidth(9, 20 * 256);      //部门
                sheet1.SetColumnWidth(10, 12 * 256);     //数量
                sheet1.SetColumnWidth(11, 12 * 256);     //金额

                sheet1.SetColumnWidth(12, 12 * 256);    //医师
                sheet1.SetColumnWidth(13, 12 * 256);    //数量
                sheet1.SetColumnWidth(14, 12 * 256);    //金额
            }
            //整理表头 
            IRow row;
            ICell cell;
            ComboBoxItem Department = (ComboBoxItem)this.comboBox1.SelectedItem;
            int f = (Department.Value == "" ? 0 : 1);

            #region 第一行
            row = sheet1.GetRow(1);
            cell = row.GetCell(0);
            cell.SetCellValue("科室：" + Department.Text + " 日期范围：" + this.dateTimePicker1.Value.ToString("yyyy-MM-dd") + " ~ " + this.dateTimePicker2.Value.ToString("yyyy-MM-dd"));
            #endregion
            #region 第二行
            row = sheet1.GetRow(2);
            cell = row.GetCell(0);
            cell.SetCellValue("库房");
            cell = row.GetCell(1);
            cell.SetCellValue("库房");
            cell = row.GetCell(2);
            cell.SetCellValue("库房");

            cell = row.GetCell(3);
            cell.SetCellValue("药品");
            cell = row.GetCell(4);
            cell.SetCellValue("药品");
            cell = row.GetCell(5);
            cell.SetCellValue("药品");
            cell = row.GetCell(6);
            cell.SetCellValue("药品");
            cell = row.GetCell(7);
            cell.SetCellValue("药品");
            cell = row.GetCell(8);
            cell.SetCellValue("药品");
            if (t == 0)
            {
                cell = row.GetCell(9);
                cell.SetCellValue("部门");
                cell = row.GetCell(10);
                cell.SetCellValue("部门");
                cell = row.GetCell(11);
                cell.SetCellValue("部门");

                cell = row.GetCell(12);
                cell.SetCellValue("医师");
                cell = row.GetCell(13);
                cell.SetCellValue("医师");
                cell = row.GetCell(14);
                cell.SetCellValue("医师");
            }
            #endregion 
            #region 第三行
            row = sheet1.GetRow(3);
            cell = row.GetCell(0);
            cell.SetCellValue("名称");
            cell = row.GetCell(1);
            cell.SetCellValue("数量");
            cell = row.GetCell(2);
            cell.SetCellValue("金额");

            cell = row.GetCell(3);
            cell.SetCellValue("药名");
            cell = row.GetCell(4);
            cell.SetCellValue("规格");
            cell = row.GetCell(5);
            cell.SetCellValue("单位");
            cell = row.GetCell(6);
            cell.SetCellValue("单价");
            cell = row.GetCell(7);
            cell.SetCellValue("数量");
            cell = row.GetCell(8);
            cell.SetCellValue("金额");
            if (t == 0)
            {
                cell = row.GetCell(9);
                cell.SetCellValue("名称");
                cell = row.GetCell(10);
                cell.SetCellValue("数量");
                cell = row.GetCell(11);
                cell.SetCellValue("金额");

                cell = row.GetCell(12);
                cell.SetCellValue("姓名");
                cell = row.GetCell(13);
                cell.SetCellValue("数量");
                cell = row.GetCell(14);
                cell.SetCellValue("金额");
            }
            #endregion 
            #region 合并表头第一行
            sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 0, 2));
            sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 3, 8));
            if (t == 0)
            {
                sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 9, 11));
                sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 12, 14));
            }
            #endregion
            #region 正文合并

            int start = 0;      //记录同组开始行号
            int end = 0;        //记录同组结束行号
            string temp = "";
            for (int j = 0; j < dt.Columns.Count - 2; j++)
            {
                start = 4;  //记录同组开始行号
                end = 4;    //记录同组结束行号

                for (int i = 0; i < m.Count; i++)
                {
                    row = sheet1.GetRow(i + 4);
                    cell = row.GetCell(j);
                    var cellText = "";
                    for (int l = 0; l < j + 1; l++)
                    {
                        cellText = cellText + row.GetCell(l).StringCellValue;
                    }
                    if (cellText == temp)       //上下行相等，记录要合并的最后一行
                    {
                        end = i + 4;
                    }
                    else//上下行不等，记录
                    {
                        if (start != end)
                        {
                            CellRangeAddress region = new CellRangeAddress(start, end, j, j);
                            sheet1.AddMergedRegion(region);
                        }
                        start = i + 4;
                        end = i + 4;
                        temp = cellText;
                    }
                }
                if (start != end)
                {
                    CellRangeAddress region = new CellRangeAddress(start, end, j, j);
                    sheet1.AddMergedRegion(region);
                }
            }

            #endregion 

            System.IO.Directory.CreateDirectory(Application.StartupPath + @"\Excel");
            string excelFile = Application.StartupPath + @"\Excel\药品库房_" + t.ToString() + f.ToString() + "_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xls";
            FileStream stream = File.OpenWrite(excelFile); ;
            workBook.Write(stream);
            stream.Close();
            MessageBox.Show("文件位置：" + excelFile, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void Excel_n1_total(List<s_n_total> n)
        {
            int t = (this.checkBox1.Checked == true ? 1 : 0);
            DataTable dt = new DataTable();
            dt.Columns.Add("科室", System.Type.GetType("System.String"));
            dt.Columns.Add("科室数量", System.Type.GetType("System.Int32"));
            dt.Columns.Add("科室金额", System.Type.GetType("System.Decimal"));
            dt.Columns.Add("项目", System.Type.GetType("System.String"));
            dt.Columns.Add("项目数量", System.Type.GetType("System.Int32"));
            dt.Columns.Add("项目金额", System.Type.GetType("System.Decimal"));
            if (t == 0)
            {
                dt.Columns.Add("部门", System.Type.GetType("System.String"));
                dt.Columns.Add("部门数量", System.Type.GetType("System.Int32"));
                dt.Columns.Add("部门金额", System.Type.GetType("System.Decimal"));
                dt.Columns.Add("医师", System.Type.GetType("System.String"));
                dt.Columns.Add("医师数量", System.Type.GetType("System.Int32"));
                dt.Columns.Add("医师金额", System.Type.GetType("System.Decimal"));
            }
            foreach (s_n_total item in n)
            {
                if (t == 0)
                {
                    dt.Rows.Add(item.科名, Convert.ToInt32(item.科码数量), item.科码金额,
                            item.项目名称, Convert.ToInt32(item.项目数量), item.项目金额,
                            item.部门名称, Convert.ToInt32(item.部门数量), item.部门金额,
                            item.医师名称, Convert.ToInt32(item.医师数量), item.医师金额);
                }
                else
                {
                    dt.Rows.Add(item.科名, Convert.ToInt32(item.科码数量), item.科码金额,
                            item.项目名称, Convert.ToInt32(item.项目数量), item.项目金额);
                }
            }

            IWorkbook workBook = new HSSFWorkbook();
            workBook = ExcelHelper.ToExcel(dt, "肥城市妇幼保健院诊疗项目(科室)分类明细");

            ISheet sheet1 = workBook.GetSheetAt(0);
            sheet1.SetColumnWidth(0, 20 * 256);     //科名
            sheet1.SetColumnWidth(1, 12 * 256);     //数量
            sheet1.SetColumnWidth(2, 12 * 256);     //金额
            sheet1.SetColumnWidth(3, 50 * 256);     //项目名称
            sheet1.SetColumnWidth(4, 12 * 256);     //数量
            sheet1.SetColumnWidth(5, 12 * 256);     //金额
            if (t == 0)
            {
                sheet1.SetColumnWidth(6, 20 * 256);     //部门
                sheet1.SetColumnWidth(7, 12 * 256);     //数量
                sheet1.SetColumnWidth(8, 12 * 256);     //金额
                sheet1.SetColumnWidth(9, 12 * 256);     //医师
                sheet1.SetColumnWidth(10, 12 * 256);     //数量
                sheet1.SetColumnWidth(11, 12 * 256);     //金额
            }

            //整理表头 
            IRow row;
            ICell cell;
            ComboBoxItem Department = (ComboBoxItem)this.comboBox1.SelectedItem;
            int f = (Department.Value == "" ? 0 : 1);

            #region 第一行
            row = sheet1.GetRow(1);
            cell = row.GetCell(0);
            cell.SetCellValue("科室：" + Department.Text + " 日期范围：" + this.dateTimePicker1.Value.ToString("yyyy-MM-dd") + " ~ " + this.dateTimePicker2.Value.ToString("yyyy-MM-dd"));
            #endregion

            #region 第二行
            row = sheet1.GetRow(2);
            cell = row.GetCell(0);
            cell.SetCellValue("科室");
            cell = row.GetCell(1);
            cell.SetCellValue("科室");
            cell = row.GetCell(2);
            cell.SetCellValue("科室");

            cell = row.GetCell(3);
            cell.SetCellValue("项目");
            cell = row.GetCell(4);
            cell.SetCellValue("项目");
            cell = row.GetCell(5);
            cell.SetCellValue("项目");
            if (t == 0)
            {
                cell = row.GetCell(6);
                cell.SetCellValue("部门");
                cell = row.GetCell(7);
                cell.SetCellValue("部门");
                cell = row.GetCell(8);
                cell.SetCellValue("部门");

                cell = row.GetCell(9);
                cell.SetCellValue("医师");
                cell = row.GetCell(10);
                cell.SetCellValue("医师");
                cell = row.GetCell(11);
                cell.SetCellValue("医师");
            }
            #endregion 
            #region 第三行
            row = sheet1.GetRow(3);
            cell = row.GetCell(0);
            cell.SetCellValue("名称");
            cell = row.GetCell(1);
            cell.SetCellValue("数量");
            cell = row.GetCell(2);
            cell.SetCellValue("金额");

            cell = row.GetCell(3);
            cell.SetCellValue("名称");
            cell = row.GetCell(4);
            cell.SetCellValue("数量");
            cell = row.GetCell(5);
            cell.SetCellValue("金额");
            if (t == 0)
            {
                cell = row.GetCell(6);
                cell.SetCellValue("名称");
                cell = row.GetCell(7);
                cell.SetCellValue("数量");
                cell = row.GetCell(8);
                cell.SetCellValue("金额");

                cell = row.GetCell(9);
                cell.SetCellValue("姓名");
                cell = row.GetCell(10);
                cell.SetCellValue("数量");
                cell = row.GetCell(11);
                cell.SetCellValue("金额");
            }
            #endregion 
            #region 合并表头第一行
            sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 0, 2));
            sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 3, 5));
            if (t == 0)
            {
                sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 6, 8));
                sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 9, 11));
            }
            #endregion

            #region 正文合并

            int start = 0;      //记录同组开始行号
            int end = 0;        //记录同组结束行号
            string temp = "";
            for (int j = 0; j < dt.Columns.Count - 2; j++)
            {
                start = 4;  //记录同组开始行号
                end = 4;    //记录同组结束行号

                for (int i = 0; i < n.Count; i++)
                {
                    row = sheet1.GetRow(i + 4);
                    cell = row.GetCell(j);
                    var cellText = "";
                    for (int l = 0; l < j + 1; l++)
                    {
                        cellText = cellText + row.GetCell(l).StringCellValue;
                    }

                    if (cellText == temp)       //上下行相等，记录要合并的最后一行
                    {
                        end = i + 4;
                    }
                    else//上下行不等，记录
                    {
                        if (start != end)
                        {
                            CellRangeAddress region = new CellRangeAddress(start, end, j, j);
                            sheet1.AddMergedRegion(region);
                        }
                        start = i + 4;
                        end = i + 4;
                        temp = cellText;
                    }
                }
                if (start != end)
                {
                    CellRangeAddress region = new CellRangeAddress(start, end, j, j);
                    sheet1.AddMergedRegion(region);
                }
            }

            #endregion 

            System.IO.Directory.CreateDirectory(Application.StartupPath + @"\Excel");
            string excelFile = Application.StartupPath + @"\Excel\诊疗科室_" + t.ToString() + f.ToString() + "_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xls";
            FileStream stream = File.OpenWrite(excelFile); ;
            workBook.Write(stream);
            stream.Close();
            MessageBox.Show("文件位置：" + excelFile, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void Excel_m2_total(List<s_m_total> m)
        {
            int t = (this.checkBox1.Checked == true ? 1 : 0);
            DataTable dt = new DataTable();
            dt.Columns.Add("类别", System.Type.GetType("System.String"));
            dt.Columns.Add("类别数量", System.Type.GetType("System.Int32"));
            dt.Columns.Add("类别金额", System.Type.GetType("System.Decimal"));
            dt.Columns.Add("药品", System.Type.GetType("System.String"));
            dt.Columns.Add("规格", System.Type.GetType("System.String"));
            dt.Columns.Add("单位", System.Type.GetType("System.String"));
            dt.Columns.Add("单价", System.Type.GetType("System.Decimal"));
            dt.Columns.Add("药品数量", System.Type.GetType("System.Int32"));
            dt.Columns.Add("药品金额", System.Type.GetType("System.Decimal"));
            if (t == 0)
            {
                dt.Columns.Add("部门", System.Type.GetType("System.String"));
                dt.Columns.Add("部门数量", System.Type.GetType("System.Int32"));
                dt.Columns.Add("部门金额", System.Type.GetType("System.Decimal"));
                dt.Columns.Add("医师", System.Type.GetType("System.String"));
                dt.Columns.Add("医师数量", System.Type.GetType("System.Int32"));
                dt.Columns.Add("医师金额", System.Type.GetType("System.Decimal"));
            }
            foreach (s_m_total item in m)
            {
                if (t == 0)
                {
                    dt.Rows.Add(item.类名, Convert.ToInt32(item.类码数量), item.类码金额,
                            item.药品名称, item.规格, item.单位, item.单价, Convert.ToInt32(item.药品数量), item.药品金额,
                            item.部门名称, Convert.ToInt32(item.部门数量), item.部门金额,
                            item.医师名称, Convert.ToInt32(item.医师数量), item.医师金额);
                }
                else
                {
                    dt.Rows.Add(item.类名, Convert.ToInt32(item.类码数量), item.类码金额,
                            item.药品名称, item.规格, item.单位, item.单价, Convert.ToInt32(item.药品数量), item.药品金额);
                }
            }
            IWorkbook workBook = new HSSFWorkbook();
            workBook = ExcelHelper.ToExcel(dt, "肥城市妇幼保健院药品销售(类别)分类明细");

            ISheet sheet1 = workBook.GetSheetAt(0);
            sheet1.SetColumnWidth(0, 20 * 256);     //仓库
            sheet1.SetColumnWidth(1, 12 * 256);     //数量
            sheet1.SetColumnWidth(2, 12 * 256);     //金额

            sheet1.SetColumnWidth(3, 30 * 256);     //药品名称
            sheet1.SetColumnWidth(4, 15 * 256);     //规格
            sheet1.SetColumnWidth(5, 6 * 256);      //单位
            sheet1.SetColumnWidth(6, 12 * 256);     //单价
            sheet1.SetColumnWidth(7, 12 * 256);     //数量
            sheet1.SetColumnWidth(8, 12 * 256);     //金额
            if (t == 0)
            {
                sheet1.SetColumnWidth(9, 20 * 256);      //部门
                sheet1.SetColumnWidth(10, 12 * 256);     //数量
                sheet1.SetColumnWidth(11, 12 * 256);     //金额

                sheet1.SetColumnWidth(12, 12 * 256);    //医师
                sheet1.SetColumnWidth(13, 12 * 256);    //数量
                sheet1.SetColumnWidth(14, 12 * 256);    //金额
            }
            //整理表头 
            IRow row;
            ICell cell;
            ComboBoxItem Department = (ComboBoxItem)this.comboBox1.SelectedItem;
            int f = (Department.Value == "" ? 0 : 1);

            #region 第一行
            row = sheet1.GetRow(1);
            cell = row.GetCell(0);
            cell.SetCellValue("科室：" + Department.Text + " 日期范围：" + this.dateTimePicker1.Value.ToString("yyyy-MM-dd") + " ~ " + this.dateTimePicker2.Value.ToString("yyyy-MM-dd"));
            #endregion
            #region 第二行
            row = sheet1.GetRow(2);
            cell = row.GetCell(0);
            cell.SetCellValue("类别");
            cell = row.GetCell(1);
            cell.SetCellValue("类别");
            cell = row.GetCell(2);
            cell.SetCellValue("类别");

            cell = row.GetCell(3);
            cell.SetCellValue("药品");
            cell = row.GetCell(4);
            cell.SetCellValue("药品");
            cell = row.GetCell(5);
            cell.SetCellValue("药品");
            cell = row.GetCell(6);
            cell.SetCellValue("药品");
            cell = row.GetCell(7);
            cell.SetCellValue("药品");
            cell = row.GetCell(8);
            cell.SetCellValue("药品");
            if (t == 0)
            {
                cell = row.GetCell(9);
                cell.SetCellValue("部门");
                cell = row.GetCell(10);
                cell.SetCellValue("部门");
                cell = row.GetCell(11);
                cell.SetCellValue("部门");

                cell = row.GetCell(12);
                cell.SetCellValue("医师");
                cell = row.GetCell(13);
                cell.SetCellValue("医师");
                cell = row.GetCell(14);
                cell.SetCellValue("医师");
            }
            #endregion 
            #region 第三行
            row = sheet1.GetRow(3);
            cell = row.GetCell(0);
            cell.SetCellValue("名称");
            cell = row.GetCell(1);
            cell.SetCellValue("数量");
            cell = row.GetCell(2);
            cell.SetCellValue("金额");

            cell = row.GetCell(3);
            cell.SetCellValue("药名");
            cell = row.GetCell(4);
            cell.SetCellValue("规格");
            cell = row.GetCell(5);
            cell.SetCellValue("单位");
            cell = row.GetCell(6);
            cell.SetCellValue("单价");
            cell = row.GetCell(7);
            cell.SetCellValue("数量");
            cell = row.GetCell(8);
            cell.SetCellValue("金额");
            if (t == 0)
            {
                cell = row.GetCell(9);
                cell.SetCellValue("名称");
                cell = row.GetCell(10);
                cell.SetCellValue("数量");
                cell = row.GetCell(11);
                cell.SetCellValue("金额");

                cell = row.GetCell(12);
                cell.SetCellValue("姓名");
                cell = row.GetCell(13);
                cell.SetCellValue("数量");
                cell = row.GetCell(14);
                cell.SetCellValue("金额");
            }
            #endregion 
            #region 合并表头第一行
            sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 0, 2));
            sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 3, 8));
            if (t == 0)
            {
                sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 9, 11));
                sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 12, 14));
            }
            #endregion
            #region 正文合并

            int start = 0;      //记录同组开始行号
            int end = 0;        //记录同组结束行号
            string temp = "";
            for (int j = 0; j < dt.Columns.Count - 2; j++)
            {
                start = 4;  //记录同组开始行号
                end = 4;    //记录同组结束行号

                for (int i = 0; i < m.Count; i++)
                {
                    row = sheet1.GetRow(i + 4);
                    cell = row.GetCell(j);
                    var cellText = "";
                    for (int l = 0; l < j + 1; l++)
                    {
                        cellText = cellText + row.GetCell(l).StringCellValue;
                    }
                    if (cellText == temp)       //上下行相等，记录要合并的最后一行
                    {
                        end = i + 4;
                    }
                    else//上下行不等，记录
                    {
                        if (start != end)
                        {
                            CellRangeAddress region = new CellRangeAddress(start, end, j, j);
                            sheet1.AddMergedRegion(region);
                        }
                        start = i + 4;
                        end = i + 4;
                        temp = cellText;
                    }
                }
                if (start != end)
                {
                    CellRangeAddress region = new CellRangeAddress(start, end, j, j);
                    sheet1.AddMergedRegion(region);
                }
            }

            #endregion 
            System.IO.Directory.CreateDirectory(Application.StartupPath + @"\Excel");
            string excelFile = Application.StartupPath + @"\Excel\药品类别_" + t.ToString() + f.ToString() + "_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xls";
            FileStream stream = File.OpenWrite(excelFile); ;
            workBook.Write(stream);
            stream.Close();
            MessageBox.Show("文件位置：" + excelFile, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void Excel_n2_total(List<s_n_total> n)
        {
            int t = (this.checkBox1.Checked == true ? 1 : 0);
            DataTable dt = new DataTable();
            dt.Columns.Add("科室", System.Type.GetType("System.String"));
            dt.Columns.Add("科码数量", System.Type.GetType("System.Int32"));
            dt.Columns.Add("科码金额", System.Type.GetType("System.Decimal"));
            dt.Columns.Add("项目", System.Type.GetType("System.String"));
            dt.Columns.Add("项目数量", System.Type.GetType("System.Int32"));
            dt.Columns.Add("项目金额", System.Type.GetType("System.Decimal"));
            if (t == 0)
            {
                dt.Columns.Add("部门", System.Type.GetType("System.String"));
                dt.Columns.Add("部门数量", System.Type.GetType("System.Int32"));
                dt.Columns.Add("部门金额", System.Type.GetType("System.Decimal"));
                dt.Columns.Add("医师", System.Type.GetType("System.String"));
                dt.Columns.Add("医师数量", System.Type.GetType("System.Int32"));
                dt.Columns.Add("医师金额", System.Type.GetType("System.Decimal"));
            }
            foreach (s_n_total item in n)
            {
                if (t == 0)
                {
                    dt.Rows.Add(item.科名, Convert.ToInt32(item.科码数量), item.科码金额,
                            item.项目名称, Convert.ToInt32(item.项目数量), item.项目金额,
                            item.部门名称, Convert.ToInt32(item.部门数量), item.部门金额,
                            item.医师名称, Convert.ToInt32(item.医师数量), item.医师金额);
                }
                else
                {
                    dt.Rows.Add(item.科名, Convert.ToInt32(item.科码数量), item.科码金额,
                            item.项目名称, Convert.ToInt32(item.项目数量), item.项目金额);
                }
            }
            IWorkbook workBook = new HSSFWorkbook();
            workBook = ExcelHelper.ToExcel(dt, "肥城市妇幼保健院诊疗项目（处方）分类明细");

            ISheet sheet1 = workBook.GetSheetAt(0);
            sheet1.SetColumnWidth(0, 20 * 256);     //科名
            sheet1.SetColumnWidth(1, 12 * 256);     //数量
            sheet1.SetColumnWidth(2, 12 * 256);     //金额
            sheet1.SetColumnWidth(3, 50 * 256);     //项目名称
            sheet1.SetColumnWidth(4, 12 * 256);     //数量
            sheet1.SetColumnWidth(5, 12 * 256);     //金额
            if (t == 0)
            {
                sheet1.SetColumnWidth(6, 20 * 256);     //部门
                sheet1.SetColumnWidth(7, 12 * 256);     //数量
                sheet1.SetColumnWidth(8, 12 * 256);     //金额
                sheet1.SetColumnWidth(9, 12 * 256);     //医师
                sheet1.SetColumnWidth(10, 12 * 256);     //数量
                sheet1.SetColumnWidth(11, 12 * 256);     //金额
            }

            //整理表头 
            IRow row;
            ICell cell;
            ComboBoxItem Department = (ComboBoxItem)this.comboBox1.SelectedItem;
            int f = (Department.Value == "" ? 0 : 1);

            #region 第一行
            row = sheet1.GetRow(1);
            cell = row.GetCell(0);
            cell.SetCellValue("科室：" + Department.Text + " 日期范围：" + this.dateTimePicker1.Value.ToString("yyyy-MM-dd") + " ~ " + this.dateTimePicker2.Value.ToString("yyyy-MM-dd"));
            #endregion
            #region 第二行
            row = sheet1.GetRow(2);
            cell = row.GetCell(0);
            cell.SetCellValue("科室");
            cell = row.GetCell(1);
            cell.SetCellValue("科室");
            cell = row.GetCell(2);
            cell.SetCellValue("科室");

            cell = row.GetCell(3);
            cell.SetCellValue("项目");
            cell = row.GetCell(4);
            cell.SetCellValue("项目");
            cell = row.GetCell(5);
            cell.SetCellValue("项目");
            if (t == 0)
            {
                cell = row.GetCell(6);
                cell.SetCellValue("部门");
                cell = row.GetCell(7);
                cell.SetCellValue("部门");
                cell = row.GetCell(8);
                cell.SetCellValue("部门");

                cell = row.GetCell(9);
                cell.SetCellValue("医师");
                cell = row.GetCell(10);
                cell.SetCellValue("医师");
                cell = row.GetCell(11);
                cell.SetCellValue("医师");
            }
            #endregion 
            #region 第三行
            row = sheet1.GetRow(3);
            cell = row.GetCell(0);
            cell.SetCellValue("名称");
            cell = row.GetCell(1);
            cell.SetCellValue("数量");
            cell = row.GetCell(2);
            cell.SetCellValue("金额");

            cell = row.GetCell(3);
            cell.SetCellValue("名称");
            cell = row.GetCell(4);
            cell.SetCellValue("数量");
            cell = row.GetCell(5);
            cell.SetCellValue("金额");
            if (t == 0)
            {
                cell = row.GetCell(6);
                cell.SetCellValue("名称");
                cell = row.GetCell(7);
                cell.SetCellValue("数量");
                cell = row.GetCell(8);
                cell.SetCellValue("金额");

                cell = row.GetCell(9);
                cell.SetCellValue("姓名");
                cell = row.GetCell(10);
                cell.SetCellValue("数量");
                cell = row.GetCell(11);
                cell.SetCellValue("金额");
            }
            #endregion 
            #region 合并表头第一行
            sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 0, 2));
            sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 3, 5));
            if (t == 0)
            {
                sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 6, 8));
                sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 9, 11));
            }
            #endregion

            #region 正文合并

            int start = 0;      //记录同组开始行号
            int end = 0;        //记录同组结束行号
            string temp = "";
            for (int j = 0; j < dt.Columns.Count - 2; j++)
            {
                start = 4;  //记录同组开始行号
                end = 4;    //记录同组结束行号

                for (int i = 0; i < n.Count; i++)
                {
                    row = sheet1.GetRow(i + 4);
                    cell = row.GetCell(j);
                    var cellText = "";
                    for (int l = 0; l < j + 1; l++)
                    {
                        cellText = cellText + row.GetCell(l).StringCellValue;
                    }

                    if (cellText == temp)       //上下行相等，记录要合并的最后一行
                    {
                        end = i + 4;
                    }
                    else//上下行不等，记录
                    {
                        if (start != end)
                        {
                            CellRangeAddress region = new CellRangeAddress(start, end, j, j);
                            sheet1.AddMergedRegion(region);
                        }
                        start = i + 4;
                        end = i + 4;
                        temp = cellText;
                    }
                }
                if (start != end)
                {
                    CellRangeAddress region = new CellRangeAddress(start, end, j, j);
                    sheet1.AddMergedRegion(region);
                }
            }

            #endregion

            System.IO.Directory.CreateDirectory(Application.StartupPath + @"\Excel");

            string excelFile = Application.StartupPath + @"\Excel\诊疗处方_" + t.ToString() + f.ToString() + "_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xls";

            FileStream stream = File.OpenWrite(excelFile); ;
            workBook.Write(stream);
            stream.Close();
            MessageBox.Show("文件位置：" + excelFile, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion


        public string GetAppSettings()
        {
            string Ret = "";
            try
            {
                Configuration config = System.Configuration.ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

                if (config.AppSettings.Settings["SQL_Name"] != null)
                {
                    MySQLServer.SQL_Name = config.AppSettings.Settings["SQL_Name"].Value;
                }
                else
                {
                    MySQLServer.SQL_Name = "";
                }

                if (config.AppSettings.Settings["SQL_ID"] != null)
                {
                    MySQLServer.SQL_ID = config.AppSettings.Settings["SQL_ID"].Value;
                }
                else
                {
                    MySQLServer.SQL_ID = "";
                }

                if (config.AppSettings.Settings["SQL_PassWord"] != null)
                {
                    MySQLServer.SQL_PassWord = config.AppSettings.Settings["SQL_PassWord"].Value;
                }
                else
                {
                    MySQLServer.SQL_PassWord = "";
                }

                if (config.AppSettings.Settings["SQL_DataBase"] != null)
                {
                    MySQLServer.SQL_DataBase = config.AppSettings.Settings["SQL_DataBase"].Value;
                }
                else
                {
                    MySQLServer.SQL_DataBase = "";
                }
            }
            catch (Exception ex)
            {

                Ret = ex.Message.ToString();
            }
            return Ret;
        }

        public string SetAppSettings()
        {
            string Ret = "";
            try
            {
                Configuration config = System.Configuration.ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);


                if (config.AppSettings.Settings["SQL_Name"] == null)
                {
                    config.AppSettings.Settings.Add("SQL_Name", MySQLServer.SQL_Name);
                }
                else
                {
                    config.AppSettings.Settings["SQL_Name"].Value = MySQLServer.SQL_Name;
                }

                if (config.AppSettings.Settings["SQL_ID"] == null)
                {
                    config.AppSettings.Settings.Add("SQL_ID", MySQLServer.SQL_ID);
                }
                else
                {
                    config.AppSettings.Settings["SQL_ID"].Value = MySQLServer.SQL_ID;
                }

                if (config.AppSettings.Settings["SQL_PassWord"] == null)
                {
                    config.AppSettings.Settings.Add("SQL_PassWord", MySQLServer.SQL_PassWord);
                }
                else
                {
                    config.AppSettings.Settings["SQL_PassWord"].Value = MySQLServer.SQL_PassWord;
                }

                if (config.AppSettings.Settings["SQL_DataBase"] == null)
                {
                    config.AppSettings.Settings.Add("SQL_DataBase", MySQLServer.SQL_DataBase);
                }
                else
                {
                    config.AppSettings.Settings["SQL_DataBase"].Value = MySQLServer.SQL_DataBase;
                }
                config.Save(ConfigurationSaveMode.Modified);
                System.Configuration.ConfigurationManager.RefreshSection("appSettings");
            }
            catch (Exception ex)
            {

                Ret = ex.Message.ToString();
            }
            return Ret;
        }

        public static SqlConnection GetSqlConnection(string SQL_Name, string SQL_DataBase, string SQL_ID, string SQL_PassWord, int Timeout = 10)
        {
            string ConnectionString = string.Format("" +
                                "Server = {0}; " +
                                "Initial Catalog = {1}; " +
                                "User ID = {2}; " +
                                "Password = {3}; " +
                                "max pool size = 800; min pool size = 300; Connect Timeout = " + Timeout,
                                 SQL_Name, SQL_DataBase, SQL_ID, SQL_PassWord);
            SqlConnection conn = new SqlConnection(ConnectionString);
            conn.Open();
            return conn;
        }

        public List<ComboBoxItem> GetAllDepartment()
        {
            List<ComboBoxItem> Department = new List<ComboBoxItem>();
            Department.Add(new ComboBoxItem { Value = "", Text = "全部" });

            if (MySQLServer.SelfConn == false)
            {
                MySQLServer.TestConnection();
            }
            if (MySQLServer.SelfConn == true)
            {
                using (var conn = GetSqlConnection(MySQLServer.SQL_Name, MySQLServer.SQL_DataBase, MySQLServer.SQL_ID, MySQLServer.SQL_PassWord))
                {
                    DBExec dbExec = new DBExec(conn, null);
                    StringBuilder strSQL = new StringBuilder();
                    strSQL.Append("SELECT  [部门编码],  [部门名称]  FROM  [H0_组织机构] Order By  [部门编码]");
                    DataTable dt = dbExec.Query(strSQL.ToString(), null, "H0_组织机构");
                    foreach (DataRow dr in dt.Rows)
                    {
                        ComboBoxItem model = new ComboBoxItem();

                        if (dr["部门编码"] != DBNull.Value)
                            model.Value = Convert.ToString(dr["部门编码"]).Trim();
                        if (dr["部门名称"] != DBNull.Value)
                            model.Text = Convert.ToString(dr["部门名称"]).Trim();
                        Department.Add(model);
                    }
                }
            }

            return Department;
        }


        #region 2018-11-16
        public List<s_m_total> Get_m1_total()
        {

            string cEnd = this.dateTimePicker2.Value.ToString("yyyy-MM-dd");
            if (DateTime.Parse(cEnd) >= DateTime.Parse("2019-01-15"))
            {
                this.dateTimePicker2.Value = DateTime.Parse("2019-01-15");
            }

            int t = (this.checkBox1.Checked == true ? 1 : 0);

            ComboBoxItem Department = (ComboBoxItem)this.comboBox1.SelectedItem;

            List<s_m_total> m_total = new List<s_m_total>();
            if (MySQLServer.SelfConn == false)
            {
                MySQLServer.TestConnection();
            }
            if (MySQLServer.SelfConn == true)
            {
                using (var conn = GetSqlConnection(MySQLServer.SQL_Name, MySQLServer.SQL_DataBase, MySQLServer.SQL_ID, MySQLServer.SQL_PassWord, 5000))
                {

                    this.toolStripStatusLabel1.Text = "正在加载存储过程，请等待......";
                    Application.DoEvents();

                    string a = this.dateTimePicker1.Value.ToString("yyyy-MM-dd");
                    string b = this.dateTimePicker2.Value.ToString("yyyy-MM-dd");
                    string c = Department.Value;
                    DataTable dt = new DataTable();
                    try
                    {
                        SqlCommand command = new SqlCommand();
                        command.Connection = conn;
                        command.CommandType = CommandType.StoredProcedure;
                        command.CommandText = "s_m1_total";
                        command.Parameters.AddRange(new SqlParameter[]{
                        new SqlParameter("@t", t),
                        new SqlParameter("@a", a),
                        new SqlParameter("@b", b),
                        new SqlParameter("@c", c)});
                        SqlDataAdapter adapter = new SqlDataAdapter();
                        adapter.SelectCommand = command;
                        adapter.Fill(dt);
                    }
                    catch (SqlException ex)
                    {
                        MessageBox.Show(ex.Message.ToString(), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        if (conn.State == ConnectionState.Open)
                        {
                            conn.Close();
                        }
                    }
                    this.toolStripStatusLabel1.Text = "正在整理数据，请等待......";
                    Application.DoEvents();
                    foreach (DataRow dr in dt.Rows)
                    {
                        s_m_total model = new s_m_total();

                        if (dr["库码"] != DBNull.Value) model.库码 = Convert.ToString(dr["库码"]).Trim();
                        if (dr["库名"] != DBNull.Value) model.库名 = Convert.ToString(dr["库名"]).Trim();

                        if (dr["sl_t2"] != DBNull.Value) model.库码数量 = Convert.ToDecimal(dr["sl_t2"]);
                        if (dr["je_t2"] != DBNull.Value) model.库码金额 = Convert.ToDecimal(dr["je_t2"]);

                        if (dr["药品编码"] != DBNull.Value) model.药品编码 = Convert.ToString(dr["药品编码"]).Trim();
                        if (dr["药品名称"] != DBNull.Value) model.药品名称 = Convert.ToString(dr["药品名称"]).Trim();
                        if (dr["规格"] != DBNull.Value) model.规格 = Convert.ToString(dr["规格"]).Trim();
                        if (dr["单位"] != DBNull.Value) model.单位 = Convert.ToString(dr["单位"]).Trim();
                        if (dr["单价"] != DBNull.Value) model.单价 = Convert.ToDecimal(dr["单价"]);
                        if (dr["sl_t3"] != DBNull.Value) model.药品数量 = Convert.ToDecimal(dr["sl_t3"]);
                        if (dr["je_t3"] != DBNull.Value) model.药品金额 = Convert.ToDecimal(dr["je_t3"]);
                        if (t == 0)
                        {
                            if (dr["部门编码"] != DBNull.Value) model.部门编码 = Convert.ToString(dr["部门编码"]).Trim();
                            if (dr["部门名称"] != DBNull.Value) model.部门名称 = Convert.ToString(dr["部门名称"]).Trim();
                            if (dr["sl_t4"] != DBNull.Value) model.部门数量 = Convert.ToDecimal(dr["sl_t4"]);
                            if (dr["je_t4"] != DBNull.Value) model.部门金额 = Convert.ToDecimal(dr["je_t4"]);
                            if (dr["医师编码"] != DBNull.Value) model.医师编码 = Convert.ToString(dr["医师编码"]).Trim();
                            if (dr["医师名称"] != DBNull.Value) model.医师名称 = Convert.ToString(dr["医师名称"]).Trim();
                            if (dr["sl_t5"] != DBNull.Value) model.医师数量 = Convert.ToDecimal(dr["sl_t5"]);
                            if (dr["je_t5"] != DBNull.Value) model.医师金额 = Convert.ToDecimal(dr["je_t5"]);
                        }
                        m_total.Add(model);
                    }
                }
            }
            else
            {
                MessageBox.Show("远程数据库连接失败！！！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return m_total;
        }

        public List<s_m_total> Get_m2_total()
        {
            string cEnd = this.dateTimePicker2.Value.ToString("yyyy-MM-dd");
            if (DateTime.Parse(cEnd) >= DateTime.Parse("2019-01-15"))
            {
                this.dateTimePicker2.Value = DateTime.Parse("2019-01-15");
            }

            int t = (this.checkBox1.Checked == true ? 1 : 0);

            ComboBoxItem Department = (ComboBoxItem)this.comboBox1.SelectedItem;

            List<s_m_total> m_total = new List<s_m_total>();
            if (MySQLServer.SelfConn == false)
            {
                MySQLServer.TestConnection();
            }
            if (MySQLServer.SelfConn == true)
            {
                using (var conn = GetSqlConnection(MySQLServer.SQL_Name, MySQLServer.SQL_DataBase, MySQLServer.SQL_ID, MySQLServer.SQL_PassWord, 5000))
                {

                    this.toolStripStatusLabel1.Text = "正在加载存储过程，请等待......";
                    Application.DoEvents();

                    string a = this.dateTimePicker1.Value.ToString("yyyy-MM-dd");
                    string b = this.dateTimePicker2.Value.ToString("yyyy-MM-dd");
                    string c = Department.Value;
                    DataTable dt = new DataTable();
                    try
                    {
                        SqlCommand command = new SqlCommand();
                        command.Connection = conn;
                        command.CommandType = CommandType.StoredProcedure;
                        command.CommandText = "s_m2_total";
                        command.Parameters.AddRange(new SqlParameter[]{
                        new SqlParameter("@t", t),
                        new SqlParameter("@a", a),
                        new SqlParameter("@b", b),
                        new SqlParameter("@c", c)});
                        SqlDataAdapter adapter = new SqlDataAdapter();
                        adapter.SelectCommand = command;
                        adapter.Fill(dt);
                    }
                    catch (SqlException ex)
                    {
                        MessageBox.Show(ex.Message.ToString(), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        if (conn.State == ConnectionState.Open)
                        {
                            conn.Close();
                        }
                    }
                    this.toolStripStatusLabel1.Text = "正在整理数据，请等待......";
                    Application.DoEvents();

                    foreach (DataRow dr in dt.Rows)
                    {
                        s_m_total model = new s_m_total();
                        if (dr["类码"] != DBNull.Value) model.类码 = Convert.ToString(dr["类码"]).Trim();
                        if (dr["类名"] != DBNull.Value) model.类名 = Convert.ToString(dr["类名"]).Trim();

                        if (dr["sl_t2"] != DBNull.Value) model.类码数量 = Convert.ToDecimal(dr["sl_t2"]);
                        if (dr["je_t2"] != DBNull.Value) model.类码金额 = Convert.ToDecimal(dr["je_t2"]);

                        if (dr["药品编码"] != DBNull.Value) model.药品编码 = Convert.ToString(dr["药品编码"]).Trim();
                        if (dr["药品名称"] != DBNull.Value) model.药品名称 = Convert.ToString(dr["药品名称"]).Trim();
                        if (dr["规格"] != DBNull.Value) model.规格 = Convert.ToString(dr["规格"]).Trim();
                        if (dr["单位"] != DBNull.Value) model.单位 = Convert.ToString(dr["单位"]).Trim();
                        if (dr["单价"] != DBNull.Value) model.单价 = Convert.ToDecimal(dr["单价"]);
                        if (dr["sl_t3"] != DBNull.Value) model.药品数量 = Convert.ToDecimal(dr["sl_t3"]);
                        if (dr["je_t3"] != DBNull.Value) model.药品金额 = Convert.ToDecimal(dr["je_t3"]);

                        if (t == 0)
                        {
                            if (dr["部门编码"] != DBNull.Value) model.部门编码 = Convert.ToString(dr["部门编码"]).Trim();
                            if (dr["部门名称"] != DBNull.Value) model.部门名称 = Convert.ToString(dr["部门名称"]).Trim();
                            if (dr["sl_t4"] != DBNull.Value) model.部门数量 = Convert.ToDecimal(dr["sl_t4"]);
                            if (dr["je_t4"] != DBNull.Value) model.部门金额 = Convert.ToDecimal(dr["je_t4"]);

                            if (dr["医师编码"] != DBNull.Value) model.医师编码 = Convert.ToString(dr["医师编码"]).Trim();
                            if (dr["医师名称"] != DBNull.Value) model.医师名称 = Convert.ToString(dr["医师名称"]).Trim();
                            if (dr["sl_t5"] != DBNull.Value) model.医师数量 = Convert.ToDecimal(dr["sl_t5"]);
                            if (dr["je_t5"] != DBNull.Value) model.医师金额 = Convert.ToDecimal(dr["je_t5"]);
                        }
                        m_total.Add(model);
                    }
                }
            }
            else
            {
                MessageBox.Show("远程数据库连接失败！！！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return m_total;
        }

        public List<s_n_total> Get_n1_total()
        {
            string cEnd = this.dateTimePicker2.Value.ToString("yyyy-MM-dd");
            if (DateTime.Parse(cEnd) >= DateTime.Parse("2019-01-15"))
            {
                this.dateTimePicker2.Value = DateTime.Parse("2019-01-15");
            }

            int t = (this.checkBox1.Checked == true ? 1 : 0);

            ComboBoxItem Department = (ComboBoxItem)this.comboBox1.SelectedItem;


            List<s_n_total> n_total = new List<s_n_total>();
            if (MySQLServer.SelfConn == false)
            {
                MySQLServer.TestConnection();
            }
            if (MySQLServer.SelfConn == true)
            {
                using (var conn = GetSqlConnection(MySQLServer.SQL_Name, MySQLServer.SQL_DataBase, MySQLServer.SQL_ID, MySQLServer.SQL_PassWord, 5000))
                {
                    this.toolStripStatusLabel1.Text = "正在加载存储过程，请等待......";
                    Application.DoEvents();
                    string a = this.dateTimePicker1.Value.ToString("yyyy-MM-dd");
                    string b = this.dateTimePicker2.Value.ToString("yyyy-MM-dd");
                    string c = Department.Value;
                    DataTable dt = new DataTable();
                    try
                    {
                        SqlCommand command = new SqlCommand();
                        command.Connection = conn;
                        command.CommandType = CommandType.StoredProcedure;
                        if (this.checkBox2.Checked == false)
                        {
                            command.CommandText = "s_n1_total";
                        }
                        else
                        {
                            command.CommandText = "s_o1_total";
                        }
                        command.Parameters.AddRange(new SqlParameter[]{
                        new SqlParameter("@t", t),
                        new SqlParameter("@a", a),
                        new SqlParameter("@b", b),
                        new SqlParameter("@c", c)});
                        SqlDataAdapter adapter = new SqlDataAdapter();
                        adapter.SelectCommand = command;
                        adapter.Fill(dt);
                    }
                    catch (SqlException ex)
                    {
                        MessageBox.Show(ex.Message.ToString(), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        if (conn.State == ConnectionState.Open)
                        {
                            conn.Close();
                        }
                    }

                    this.toolStripStatusLabel1.Text = "正在整理数据，请等待......";
                    Application.DoEvents();

                    foreach (DataRow dr in dt.Rows)
                    {
                        s_n_total model = new s_n_total();
                        if (dr["科码"] != DBNull.Value) model.科码 = Convert.ToString(dr["科码"]).Trim();
                        if (dr["科名"] != DBNull.Value) model.科名 = Convert.ToString(dr["科名"]).Trim();

                        if (dr["科码数量"] != DBNull.Value) model.科码数量 = Convert.ToDecimal(dr["科码数量"]);
                        if (dr["科码金额"] != DBNull.Value) model.科码金额 = Convert.ToDecimal(dr["科码金额"]);


                        if (this.checkBox2.Checked == false)
                        {
                            if (dr["项目编码"] != DBNull.Value) model.项目编码 = Convert.ToString(dr["项目编码"]).Trim();
                            if (dr["项目名称"] != DBNull.Value) model.项目名称 = Convert.ToString(dr["项目名称"]).Trim();
                        }
                        else
                        {
                            if (dr["收入编码"] != DBNull.Value) model.项目编码 = Convert.ToString(dr["收入编码"]).Trim();
                            if (dr["收入类型"] != DBNull.Value) model.项目名称 = Convert.ToString(dr["收入类型"]).Trim();
                        }

                        if (dr["项目数量"] != DBNull.Value) model.项目数量 = Convert.ToDecimal(dr["项目数量"]);
                        if (dr["项目金额"] != DBNull.Value) model.项目金额 = Convert.ToDecimal(dr["项目金额"]);
                        if (t == 0)
                        {
                            if (dr["部门编码"] != DBNull.Value) model.部门编码 = Convert.ToString(dr["部门编码"]).Trim();
                            if (dr["部门名称"] != DBNull.Value) model.部门名称 = Convert.ToString(dr["部门名称"]).Trim();
                            if (dr["部门数量"] != DBNull.Value) model.部门数量 = Convert.ToDecimal(dr["部门数量"]);
                            if (dr["部门金额"] != DBNull.Value) model.部门金额 = Convert.ToDecimal(dr["部门金额"]);

                            if (dr["医师编码"] != DBNull.Value) model.医师编码 = Convert.ToString(dr["医师编码"]).Trim();
                            if (dr["医师名称"] != DBNull.Value) model.医师名称 = Convert.ToString(dr["医师名称"]).Trim();
                            if (dr["医师数量"] != DBNull.Value) model.医师数量 = Convert.ToDecimal(dr["医师数量"]);
                            if (dr["医师金额"] != DBNull.Value) model.医师金额 = Convert.ToDecimal(dr["医师金额"]);
                        }
                        n_total.Add(model);
                    }
                }
            }
            else
            {
                MessageBox.Show("远程数据库连接失败！！！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return n_total;

        }

        public List<s_n_total> Get_n2_total()
        {
            string cEnd = this.dateTimePicker2.Value.ToString("yyyy-MM-dd");
            if (DateTime.Parse(cEnd) >= DateTime.Parse("2019-01-15"))
            {
                this.dateTimePicker2.Value = DateTime.Parse("2019-01-15");
            }

            int t = (this.checkBox1.Checked == true ? 1 : 0);

            ComboBoxItem Department = (ComboBoxItem)this.comboBox1.SelectedItem;


            List<s_n_total> n_total = new List<s_n_total>();
            if (MySQLServer.SelfConn == false)
            {
                MySQLServer.TestConnection();
            }
            if (MySQLServer.SelfConn == true)
            {
                using (var conn = GetSqlConnection(MySQLServer.SQL_Name, MySQLServer.SQL_DataBase, MySQLServer.SQL_ID, MySQLServer.SQL_PassWord, 5000))
                {
                    this.toolStripStatusLabel1.Text = "正在加载存储过程，请等待......";
                    Application.DoEvents();

                    string a = this.dateTimePicker1.Value.ToString("yyyy-MM-dd");
                    string b = this.dateTimePicker2.Value.ToString("yyyy-MM-dd");
                    string c = Department.Value;
                    DataTable dt = new DataTable();
                    try
                    {
                        SqlCommand command = new SqlCommand();
                        command.Connection = conn;
                        command.CommandType = CommandType.StoredProcedure;
                        if (this.checkBox2.Checked == false)
                        {
                            command.CommandText = "s_n2_total";
                        }
                        else
                        {
                            command.CommandText = "s_o2_total";
                        }
                        command.Parameters.AddRange(new SqlParameter[]{
                        new SqlParameter("@t", t),
                        new SqlParameter("@a", a),
                        new SqlParameter("@b", b),
                        new SqlParameter("@c", c)});
                        SqlDataAdapter adapter = new SqlDataAdapter();
                        adapter.SelectCommand = command;
                        adapter.Fill(dt);
                    }
                    catch (SqlException ex)
                    {
                        MessageBox.Show(ex.Message.ToString(), "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        if (conn.State == ConnectionState.Open)
                        {
                            conn.Close();
                        }
                    }
                    this.toolStripStatusLabel1.Text = "正在整理数据，请等待......";
                    Application.DoEvents();

                    foreach (DataRow dr in dt.Rows)
                    {
                        s_n_total model = new s_n_total();
                        if (dr["科码"] != DBNull.Value) model.科码 = Convert.ToString(dr["科码"]).Trim();
                        if (dr["科名"] != DBNull.Value) model.科名 = Convert.ToString(dr["科名"]).Trim();
                        if (dr["科码数量"] != DBNull.Value) model.科码数量 = Convert.ToDecimal(dr["科码数量"]);
                        if (dr["科码金额"] != DBNull.Value) model.科码金额 = Convert.ToDecimal(dr["科码金额"]);

                        if (this.checkBox2.Checked == false)
                        {
                            if (dr["项目编码"] != DBNull.Value) model.项目编码 = Convert.ToString(dr["项目编码"]).Trim();
                            if (dr["项目名称"] != DBNull.Value) model.项目名称 = Convert.ToString(dr["项目名称"]).Trim();
                        }
                        else
                        {
                            if (dr["收入编码"] != DBNull.Value) model.项目编码 = Convert.ToString(dr["收入编码"]).Trim();
                            if (dr["收入类型"] != DBNull.Value) model.项目名称 = Convert.ToString(dr["收入类型"]).Trim();
                        }

                        if (dr["项目数量"] != DBNull.Value) model.项目数量 = Convert.ToDecimal(dr["项目数量"]);
                        if (dr["项目金额"] != DBNull.Value) model.项目金额 = Convert.ToDecimal(dr["项目金额"]);
                        if (t == 0)
                        {
                            if (dr["部门编码"] != DBNull.Value) model.部门编码 = Convert.ToString(dr["部门编码"]).Trim();
                            if (dr["部门名称"] != DBNull.Value) model.部门名称 = Convert.ToString(dr["部门名称"]).Trim();
                            if (dr["部门数量"] != DBNull.Value) model.部门数量 = Convert.ToDecimal(dr["部门数量"]);
                            if (dr["部门金额"] != DBNull.Value) model.部门金额 = Convert.ToDecimal(dr["部门金额"]);

                            if (dr["医师编码"] != DBNull.Value) model.医师编码 = Convert.ToString(dr["医师编码"]).Trim();
                            if (dr["医师名称"] != DBNull.Value) model.医师名称 = Convert.ToString(dr["医师名称"]).Trim();
                            if (dr["医师数量"] != DBNull.Value) model.医师数量 = Convert.ToDecimal(dr["医师数量"]);
                            if (dr["医师金额"] != DBNull.Value) model.医师金额 = Convert.ToDecimal(dr["医师金额"]);
                        }
                        n_total.Add(model);
                    }
                }
            }
            else
            {
                MessageBox.Show("远程数据库连接失败！！！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return n_total;

        }

        #endregion
        private void button6_Click(object sender, EventArgs e)
        {
            if (MySQLServer.SelfConn == true)
            {
                this.comboBox1.Items.Clear();
                List<ComboBoxItem> Departments = GetAllDepartment();
                foreach (ComboBoxItem Department in Departments)
                {
                    this.comboBox1.Items.Add(Department);
                }
                if (this.comboBox1.Items.Count > 0) { this.comboBox1.SelectedIndex = 0; }

            }
            else
            {
                MessageBox.Show("远程数据库连接失败！！！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (this.radioButton1.Checked == true)
            {
                this.radioButton3.Checked = true;
                this.radioButton3.Enabled = false;
                this.radioButton4.Enabled = false;
                this.radioButton5.Enabled = false;
                this.radioButton6.Enabled = false;
                this.radioButton7.Enabled = false;
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (this.radioButton2.Checked == true)
            {
                this.radioButton3.Checked = true;
                this.radioButton3.Enabled = true;
                this.radioButton4.Enabled = true;
                this.radioButton5.Enabled = true;
                this.radioButton6.Enabled = true;
                this.radioButton7.Enabled = true;
            }
        }


        private void Get_nx_total(int nType)
        {
            DataTable dt = new DataTable();
            DataTable dt1 = new DataTable();
            DataTable dt2 = new DataTable();
            DataTable dt3 = new DataTable();

            #region 2018-12-03
            string cEnd = this.dateTimePicker2.Value.ToString("yyyy-MM-dd");
            if (DateTime.Parse(cEnd) >= DateTime.Parse("2019-01-15"))
            {
                this.dateTimePicker2.Value = DateTime.Parse("2019-01-15");
            }

            ComboBoxItem Department = (ComboBoxItem)this.comboBox1.SelectedItem;

            if (MySQLServer.SelfConn == false)
            {
                MySQLServer.TestConnection();
            }
            if (MySQLServer.SelfConn == true)
            {
                using (var conn = GetSqlConnection(MySQLServer.SQL_Name, MySQLServer.SQL_DataBase, MySQLServer.SQL_ID, MySQLServer.SQL_PassWord, 5000))
                {

                    this.toolStripStatusLabel1.Text = "正在加载存储过程，请等待......";
                    Application.DoEvents();


                    string a = this.dateTimePicker1.Value.ToString("yyyy-MM-dd");
                    string b = this.dateTimePicker2.Value.ToString("yyyy-MM-dd");
                    string c = Department.Value;

                    UpdateDefault(conn);

                    //==================
                    string cSql0 = "";
                    string cSql1 = "";
                    string cSql2 = "";
                    string cSql3 = "";
                    string cWh01 = " 日期>='" + a + "' And 日期<='" + b + "'";
                    if (c != "")
                    {
                        if (this.comboBox3.SelectedIndex == 0)
                        {
                            cWh01 = cWh01 + " And 科码='" + c + "'";
                        }
                        else
                        {
                            cWh01 = cWh01 + " And 部门编码='" + c + "'";
                        }
                    }

                    if (nType == 1)
                    {
                        //=========纵向
                        if (this.comboBox3.SelectedIndex == 0)
                        {
                            cSql1 = " Select 科码 , 科名 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 科码 , 科名 order by 科码 , 科名 ";
                        }
                        else
                        {
                            cSql1 = " Select 部门编码 as 科码 , 部门名称 as 科名 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 部门编码 , 部门名称 order by 部门编码 , 部门名称";
                        }
                        //==========横向 
                        if (checkBox2.Checked == true)
                        {
                            cSql2 = " Select 收入编码 as 项目编码 , 收入类型 as 项目名称 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 收入编码 , 收入类型 order by 收入编码 , 收入类型 ";
                        }
                        else
                        {
                            cSql2 = " Select 项目编码  , 项目名称 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 项目编码 , 项目名称 order by 项目编码 , 项目名称 ";
                        }
                    }
                    else if (nType == 2)
                    {
                        //=========纵向
                        if (checkBox2.Checked == true)
                        {
                            cSql1 = " Select 收入编码 as 项目编码 , 收入类型 as 项目名称 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 收入编码 , 收入类型 order by 收入编码 , 收入类型 ";
                        }
                        else
                        {
                            cSql1 = " Select 项目编码  , 项目名称 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 项目编码 , 项目名称 order by 项目编码 , 项目名称 ";
                        }
                        //==========横向 
                        if (this.comboBox3.SelectedIndex == 0)
                        {
                            cSql2 = " Select 科码 , 科名 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 科码 , 科名 order by 科码 , 科名 ";
                        }
                        else
                        {
                            cSql2 = " Select 部门编码 as 科码 , 部门名称 as 科名 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 部门编码 , 部门名称 order by 部门编码 , 部门名称 ";
                        }

                    }
                    else if (nType == 3)
                    {
                        //=========纵向
                        cSql1 = " Select 医师编码 , 医师名称 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 医师编码 , 医师名称 order by 医师编码 , 医师名称 ";
                        //==========横向 
                        if (checkBox2.Checked == true)
                        {
                            cSql2 = " Select 收入编码 as 项目编码 , 收入类型 as 项目名称 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 收入编码 , 收入类型 order by 收入编码 , 收入类型 ";
                        }
                        else
                        {
                            cSql2 = " Select 项目编码  , 项目名称 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 项目编码 , 项目名称 order by 项目编码 , 项目名称 ";
                        }
                    }
                    else if (nType == 4)
                    {
                        //=========纵向
                        if (checkBox2.Checked == true)
                        {
                            cSql1 = " Select 收入编码 as 项目编码 , 收入类型 as 项目名称 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 收入编码 , 收入类型 order by 收入编码 , 收入类型 ";
                        }
                        else
                        {
                            cSql1 = " Select 项目编码  , 项目名称 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 项目编码 , 项目名称 order by 项目编码 , 项目名称 ";
                        }
                        //==========横向 
                        cSql2 = " Select 医师编码 , 医师名称 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 医师编码 , 医师名称 order by 医师编码 , 医师名称 ";

                    }

                    this.toolStripStatusLabel1.Text = "正在归集数据，请等待......";
                    Application.DoEvents();

                    dt1 = ExecuteDataTable(conn, cSql1);


                    dt2 = ExecuteDataTable(conn, cSql2);


                    cSql3 = " Select 科码 , 科名 , 部门编码 , 部门名称 ,收入编码 , 收入类型 ,项目编码 , 项目名称 ,医师编码 , 医师名称 , 金额 from H4_收款记录 Where " + cWh01;
                    dt3 = ExecuteDataTable(conn, cSql3);

                    //================== 
                    dt.Columns.Add("code", System.Type.GetType("System.String"));
                    dt.Columns.Add("name", System.Type.GetType("System.String"));

                    string[] Cols = new string[dt2.Rows.Count];
                    int nCol = 0;
                    foreach (DataRow dr2 in dt2.Rows)
                    {
                        string ColumnName = "";
                        if (nType == 1)
                        {
                            ColumnName = "col_" + Convert.ToString(dr2["项目编码"]).Trim();
                            Cols[nCol] = Convert.ToString(dr2["项目编码"]).Trim();
                        }
                        else if (nType == 2)
                        {
                            ColumnName = "col_" + Convert.ToString(dr2["科码"]).Trim();
                            Cols[nCol] = Convert.ToString(dr2["科码"]).Trim();
                        }
                        else if (nType == 3)
                        {
                            ColumnName = "col_" + Convert.ToString(dr2["项目编码"]).Trim();
                            Cols[nCol] = Convert.ToString(dr2["项目编码"]).Trim();
                        }
                        else if (nType == 4)
                        {
                            ColumnName = "col_" + Convert.ToString(dr2["医师编码"]).Trim();
                            Cols[nCol] = Convert.ToString(dr2["医师编码"]).Trim();
                        }
                        dt.Columns.Add(ColumnName, System.Type.GetType("System.Decimal"));
                        nCol = nCol + 1;
                    }

                    string[] Rows = new string[dt1.Rows.Count];
                    int nRow = 0;
                    foreach (DataRow dr1 in dt1.Rows)
                    {
                        string code = "";
                        string name = "";
                        if (nType == 1)
                        {
                            code = Convert.ToString(dr1["科码"]).Trim();
                            name = Convert.ToString(dr1["科名"]).Trim();
                        }
                        else if (nType == 2)
                        {
                            code = Convert.ToString(dr1["项目编码"]).Trim();
                            name = Convert.ToString(dr1["项目名称"]).Trim();
                        }
                        else if (nType == 3)
                        {
                            code = Convert.ToString(dr1["医师编码"]).Trim();
                            name = Convert.ToString(dr1["医师名称"]).Trim();
                        }
                        else if (nType == 4)
                        {
                            code = Convert.ToString(dr1["项目编码"]).Trim();
                            name = Convert.ToString(dr1["项目名称"]).Trim();
                        }


                        dt.Rows.Add(code, name);
                        Rows[nRow] = code;
                        nRow = nRow + 1;
                    }

                    this.toolStripStatusLabel1.Text = "正在整理数据，请等待......";
                    Application.DoEvents();

                    foreach (DataRow dr3 in dt3.Rows)
                    {
                        string cRma = "";
                        string cRmg = "";
                        string cCma = "";
                        string cCmg = "";

                        decimal je = 0;
                        decimal value = 0;


                        if (nType == 1)
                        {
                            if (this.comboBox3.SelectedIndex == 0)
                            {
                                cRma = Convert.ToString(dr3["科码"]).Trim();
                                cRmg = Convert.ToString(dr3["科名"]).Trim();
                            }
                            else
                            {
                                cRma = Convert.ToString(dr3["部门编码"]).Trim();
                                cRmg = Convert.ToString(dr3["部门名称"]).Trim();
                            }

                            if (checkBox2.Checked == true)
                            {
                                cCma = Convert.ToString(dr3["收入编码"]).Trim();
                                cCmg = Convert.ToString(dr3["收入类型"]).Trim();
                            }
                            else
                            {
                                cCma = Convert.ToString(dr3["项目编码"]).Trim();
                                cCmg = Convert.ToString(dr3["项目名称"]).Trim();
                            }
                        }
                        else if (nType == 2)
                        {


                            if (checkBox2.Checked == true)
                            {
                                cRma = Convert.ToString(dr3["收入编码"]).Trim();
                                cRmg = Convert.ToString(dr3["收入类型"]).Trim();
                            }
                            else
                            {
                                cRma = Convert.ToString(dr3["项目编码"]).Trim();
                                cRmg = Convert.ToString(dr3["项目名称"]).Trim();
                            }

                            if (this.comboBox3.SelectedIndex == 0)
                            {
                                cCma = Convert.ToString(dr3["科码"]).Trim();
                                cCmg = Convert.ToString(dr3["科名"]).Trim();
                            }
                            else
                            {
                                cCma = Convert.ToString(dr3["部门编码"]).Trim();
                                cCmg = Convert.ToString(dr3["部门名称"]).Trim();
                            }
                        }
                        else if (nType == 3)
                        {
                            cRma = Convert.ToString(dr3["医师编码"]).Trim();
                            cRmg = Convert.ToString(dr3["医师名称"]).Trim();

                            if (checkBox2.Checked == true)
                            {
                                cCma = Convert.ToString(dr3["收入编码"]).Trim();
                                cCmg = Convert.ToString(dr3["收入类型"]).Trim();
                            }
                            else
                            {
                                cCma = Convert.ToString(dr3["项目编码"]).Trim();
                                cCmg = Convert.ToString(dr3["项目名称"]).Trim();
                            }
                        }
                        else if (nType == 4)
                        {
                            if (checkBox2.Checked == true)
                            {
                                cRma = Convert.ToString(dr3["收入编码"]).Trim();
                                cRmg = Convert.ToString(dr3["收入类型"]).Trim();
                            }
                            else
                            {
                                cRma = Convert.ToString(dr3["项目编码"]).Trim();
                                cRmg = Convert.ToString(dr3["项目名称"]).Trim();
                            }
                            cCma = Convert.ToString(dr3["医师编码"]).Trim();
                            cCmg = Convert.ToString(dr3["医师名称"]).Trim();
                        }

                        je = Convert.ToDecimal(dr3["金额"]);

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            DataRow dr = dt.Rows[i];
                            string code = Convert.ToString(dr["code"]).Trim();
                            if (cRma == code)
                            {
                                for (int j = 0; j < Cols.Length; j++)
                                {
                                    string ColName = Cols[j];
                                    if (ColName == cCma)
                                    {

                                        if (dr["col_" + ColName] != DBNull.Value)
                                        {
                                            value = Convert.ToDecimal(dr["col_" + ColName]);
                                        }
                                        else
                                        {
                                            value = 0;
                                        }
                                        value = value + je;
                                        dr["col_" + ColName] = value;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("远程数据库连接失败！！！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            #endregion
            this.dataGridView1.DataSource = null;
            this.dataGridView1.DataSource = dt;
            if (nType == 1)
            {
                this.dataGridView1.Columns[0].HeaderText = "科室编码";
                this.dataGridView1.Columns[1].HeaderText = "科室名称";
            }
            else if (nType == 2)
            {
                this.dataGridView1.Columns[0].HeaderText = "项目编码";
                this.dataGridView1.Columns[1].HeaderText = "项目名称";
            }
            else if (nType == 3)
            {
                this.dataGridView1.Columns[0].HeaderText = "医师编码";
                this.dataGridView1.Columns[1].HeaderText = "医师名称";
            }
            else if (nType == 4)
            {
                this.dataGridView1.Columns[0].HeaderText = "项目编码";
                this.dataGridView1.Columns[1].HeaderText = "项目名称";
            }



            foreach (DataRow dr2 in dt2.Rows)
            {
                string ColumnName = "";
                string HeaderText = "";
                if (nType == 1)
                {
                    ColumnName = "col_" + Convert.ToString(dr2["项目编码"]).Trim();
                    HeaderText = Convert.ToString(dr2["项目名称"]).Trim();
                }
                else if (nType == 2)
                {
                    ColumnName = "col_" + Convert.ToString(dr2["科码"]).Trim();
                    HeaderText = Convert.ToString(dr2["科名"]).Trim();
                }
                else if (nType == 3)
                {
                    ColumnName = "col_" + Convert.ToString(dr2["项目编码"]).Trim();
                    HeaderText = Convert.ToString(dr2["项目名称"]).Trim();
                }
                else if (nType == 4)
                {
                    ColumnName = "col_" + Convert.ToString(dr2["医师编码"]).Trim();
                    HeaderText = Convert.ToString(dr2["医师名称"]).Trim();
                }

                this.dataGridView1.Columns[ColumnName].HeaderText = HeaderText;
                this.dataGridView1.Columns[ColumnName].DefaultCellStyle.Format = "0.00";
                this.dataGridView1.Columns[ColumnName].DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;

            }


            MessageBox.Show("数据获取完毕！！！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }



        private void Excel_nx_total(int nType)
        {
            DataTable dt = new DataTable();
            DataTable dt1 = new DataTable();
            DataTable dt2 = new DataTable();
            DataTable dt3 = new DataTable();

            #region 2018-12-03
            string cEnd = this.dateTimePicker2.Value.ToString("yyyy-MM-dd");
            if (DateTime.Parse(cEnd) >= DateTime.Parse("2019-01-15"))
            {
                this.dateTimePicker2.Value = DateTime.Parse("2019-01-15");
            }

            ComboBoxItem Department = (ComboBoxItem)this.comboBox1.SelectedItem;

            if (MySQLServer.SelfConn == false)
            {
                MySQLServer.TestConnection();
            }
            if (MySQLServer.SelfConn == true)
            {
                using (var conn = GetSqlConnection(MySQLServer.SQL_Name, MySQLServer.SQL_DataBase, MySQLServer.SQL_ID, MySQLServer.SQL_PassWord, 5000))
                {

                    this.toolStripStatusLabel1.Text = "正在加载存储过程，请等待......";
                    Application.DoEvents();


                    string a = this.dateTimePicker1.Value.ToString("yyyy-MM-dd");
                    string b = this.dateTimePicker2.Value.ToString("yyyy-MM-dd");
                    string c = Department.Value;

                    UpdateDefault(conn);

                    //==================
                    string cSql0 = "";
                    string cSql1 = "";
                    string cSql2 = "";
                    string cSql3 = "";
                    string cWh01 = " 日期>='" + a + "' And 日期<='" + b + "'";
                    if (c != "")
                    {
                        if (this.comboBox3.SelectedIndex == 0)
                        {
                            cWh01 = cWh01 + " And 科码='" + c + "'";
                        }
                        else
                        {
                            cWh01 = cWh01 + " And 部门编码='" + c + "'";
                        }
                    }

                    if (nType == 1)
                    {
                        //=========纵向
                        if (this.comboBox3.SelectedIndex == 0)
                        {
                            cSql1 = " Select 科码 , 科名 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 科码 , 科名 order by 科码 , 科名 ";
                        }
                        else
                        {
                            cSql1 = " Select 部门编码 as 科码 , 部门名称 as 科名 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 部门编码 , 部门名称 order by 部门编码 , 部门名称";
                        }
                        //==========横向 
                        if (checkBox2.Checked == true)
                        {
                            cSql2 = " Select 收入编码 as 项目编码 , 收入类型 as 项目名称 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 收入编码 , 收入类型 order by 收入编码 , 收入类型 ";
                        }
                        else
                        {
                            cSql2 = " Select 项目编码  , 项目名称 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 项目编码 , 项目名称 order by 项目编码 , 项目名称 ";
                        }
                    }
                    else if (nType == 2)
                    {
                        //=========纵向
                        if (checkBox2.Checked == true)
                        {
                            cSql1 = " Select 收入编码 as 项目编码 , 收入类型 as 项目名称 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 收入编码 , 收入类型 order by 收入编码 , 收入类型 ";
                        }
                        else
                        {
                            cSql1 = " Select 项目编码  , 项目名称 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 项目编码 , 项目名称 order by 项目编码 , 项目名称 ";
                        }
                        //==========横向 
                        if (this.comboBox3.SelectedIndex == 0)
                        {
                            cSql2 = " Select 科码 , 科名 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 科码 , 科名 order by 科码 , 科名 ";
                        }
                        else
                        {
                            cSql2 = " Select 部门编码 as 科码 , 部门名称 as 科名 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 部门编码 , 部门名称 order by 部门编码 , 部门名称 ";
                        }

                    }
                    else if (nType == 3)
                    {
                        //=========纵向
                        cSql1 = " Select 医师编码 , 医师名称 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 医师编码 , 医师名称 order by 医师编码 , 医师名称 ";
                        //==========横向 
                        if (checkBox2.Checked == true)
                        {
                            cSql2 = " Select 收入编码 as 项目编码 , 收入类型 as 项目名称 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 收入编码 , 收入类型 order by 收入编码 , 收入类型 ";
                        }
                        else
                        {
                            cSql2 = " Select 项目编码  , 项目名称 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 项目编码 , 项目名称 order by 项目编码 , 项目名称 ";
                        }
                    }
                    else if (nType == 4)
                    {
                        //=========纵向
                        if (checkBox2.Checked == true)
                        {
                            cSql1 = " Select 收入编码 as 项目编码 , 收入类型 as 项目名称 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 收入编码 , 收入类型 order by 收入编码 , 收入类型 ";
                        }
                        else
                        {
                            cSql1 = " Select 项目编码  , 项目名称 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 项目编码 , 项目名称 order by 项目编码 , 项目名称 ";
                        }
                        //==========横向 
                        cSql2 = " Select 医师编码 , 医师名称 , sum(金额) as 金额 from H4_收款记录 Where " + cWh01 + " Group by 医师编码 , 医师名称 order by 医师编码 , 医师名称 ";

                    }

                    this.toolStripStatusLabel1.Text = "正在归集数据，请等待......";
                    Application.DoEvents();

                    dt1 = ExecuteDataTable(conn, cSql1);


                    dt2 = ExecuteDataTable(conn, cSql2);


                    cSql3 = " Select 科码 , 科名 , 部门编码 , 部门名称 ,收入编码 , 收入类型 ,项目编码 , 项目名称 ,医师编码 , 医师名称 , 金额 from H4_收款记录 Where " + cWh01;
                    dt3 = ExecuteDataTable(conn, cSql3);

                    //================== 
                    dt.Columns.Add("code", System.Type.GetType("System.String"));
                    dt.Columns.Add("name", System.Type.GetType("System.String"));

                    string[] Cols = new string[dt2.Rows.Count];
                    int nCol = 0;
                    foreach (DataRow dr2 in dt2.Rows)
                    {
                        string ColumnName = "";
                        if (nType == 1)
                        {
                            ColumnName = "col_" + Convert.ToString(dr2["项目编码"]).Trim();
                            Cols[nCol] = Convert.ToString(dr2["项目编码"]).Trim();
                        }
                        else if (nType == 2)
                        {
                            ColumnName = "col_" + Convert.ToString(dr2["科码"]).Trim();
                            Cols[nCol] = Convert.ToString(dr2["科码"]).Trim();
                        }
                        else if (nType == 3)
                        {
                            ColumnName = "col_" + Convert.ToString(dr2["项目编码"]).Trim();
                            Cols[nCol] = Convert.ToString(dr2["项目编码"]).Trim();
                        }
                        else if (nType == 4)
                        {
                            ColumnName = "col_" + Convert.ToString(dr2["医师编码"]).Trim();
                            Cols[nCol] = Convert.ToString(dr2["医师编码"]).Trim();
                        }
                        dt.Columns.Add(ColumnName, System.Type.GetType("System.Decimal"));
                        nCol = nCol + 1;
                    }

                    string[] Rows = new string[dt1.Rows.Count];
                    int nRow = 0;
                    foreach (DataRow dr1 in dt1.Rows)
                    {
                        string code = "";
                        string name = "";
                        if (nType == 1)
                        {
                            code = Convert.ToString(dr1["科码"]).Trim();
                            name = Convert.ToString(dr1["科名"]).Trim();
                        }
                        else if (nType == 2)
                        {
                            code = Convert.ToString(dr1["项目编码"]).Trim();
                            name = Convert.ToString(dr1["项目名称"]).Trim();
                        }
                        else if (nType == 3)
                        {
                            code = Convert.ToString(dr1["医师编码"]).Trim();
                            name = Convert.ToString(dr1["医师名称"]).Trim();
                        }
                        else if (nType == 4)
                        {
                            code = Convert.ToString(dr1["项目编码"]).Trim();
                            name = Convert.ToString(dr1["项目名称"]).Trim();
                        }


                        dt.Rows.Add(code, name);
                        Rows[nRow] = code;
                        nRow = nRow + 1;
                    }

                    this.toolStripStatusLabel1.Text = "正在整理数据，请等待......";
                    Application.DoEvents();

                    foreach (DataRow dr3 in dt3.Rows)
                    {
                        string cRma = "";
                        string cRmg = "";
                        string cCma = "";
                        string cCmg = "";

                        decimal je = 0;
                        decimal value = 0;


                        if (nType == 1)
                        {
                            if (this.comboBox3.SelectedIndex == 0)
                            {
                                cRma = Convert.ToString(dr3["科码"]).Trim();
                                cRmg = Convert.ToString(dr3["科名"]).Trim();
                            }
                            else
                            {
                                cRma = Convert.ToString(dr3["部门编码"]).Trim();
                                cRmg = Convert.ToString(dr3["部门名称"]).Trim();
                            }

                            if (checkBox2.Checked == true)
                            {
                                cCma = Convert.ToString(dr3["收入编码"]).Trim();
                                cCmg = Convert.ToString(dr3["收入类型"]).Trim();
                            }
                            else
                            {
                                cCma = Convert.ToString(dr3["项目编码"]).Trim();
                                cCmg = Convert.ToString(dr3["项目名称"]).Trim();
                            }
                        }
                        else if (nType == 2)
                        {


                            if (checkBox2.Checked == true)
                            {
                                cRma = Convert.ToString(dr3["收入编码"]).Trim();
                                cRmg = Convert.ToString(dr3["收入类型"]).Trim();
                            }
                            else
                            {
                                cRma = Convert.ToString(dr3["项目编码"]).Trim();
                                cRmg = Convert.ToString(dr3["项目名称"]).Trim();
                            }

                            if (this.comboBox3.SelectedIndex == 0)
                            {
                                cCma = Convert.ToString(dr3["科码"]).Trim();
                                cCmg = Convert.ToString(dr3["科名"]).Trim();
                            }
                            else
                            {
                                cCma = Convert.ToString(dr3["部门编码"]).Trim();
                                cCmg = Convert.ToString(dr3["部门名称"]).Trim();
                            }
                        }
                        else if (nType == 3)
                        {
                            cRma = Convert.ToString(dr3["医师编码"]).Trim();
                            cRmg = Convert.ToString(dr3["医师名称"]).Trim();

                            if (checkBox2.Checked == true)
                            {
                                cCma = Convert.ToString(dr3["收入编码"]).Trim();
                                cCmg = Convert.ToString(dr3["收入类型"]).Trim();
                            }
                            else
                            {
                                cCma = Convert.ToString(dr3["项目编码"]).Trim();
                                cCmg = Convert.ToString(dr3["项目名称"]).Trim();
                            }
                        }
                        else if (nType == 4)
                        {
                            if (checkBox2.Checked == true)
                            {
                                cRma = Convert.ToString(dr3["收入编码"]).Trim();
                                cRmg = Convert.ToString(dr3["收入类型"]).Trim();
                            }
                            else
                            {
                                cRma = Convert.ToString(dr3["项目编码"]).Trim();
                                cRmg = Convert.ToString(dr3["项目名称"]).Trim();
                            }
                            cCma = Convert.ToString(dr3["医师编码"]).Trim();
                            cCmg = Convert.ToString(dr3["医师名称"]).Trim();
                        }

                        je = Convert.ToDecimal(dr3["金额"]);

                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            DataRow dr = dt.Rows[i];
                            string code = Convert.ToString(dr["code"]).Trim();
                            if (cRma == code)
                            {
                                for (int j = 0; j < Cols.Length; j++)
                                {
                                    string ColName = Cols[j];
                                    if (ColName == cCma)
                                    {

                                        if (dr["col_" + ColName] != DBNull.Value)
                                        {
                                            value = Convert.ToDecimal(dr["col_" + ColName]);
                                        }
                                        else
                                        {
                                            value = 0;
                                        }
                                        value = value + je;
                                        dr["col_" + ColName] = value;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    this.toolStripStatusLabel1.Text = "导出Excel，请等待......";
                    Application.DoEvents();
                    #region 


                    IWorkbook workBook = new HSSFWorkbook();

                    if (nType == 1)
                    {
                        workBook = ExcelHelper.ToExcelPro(dt, "肥城市妇幼保健院诊疗项目（科室项目)表");
                    }
                    else if (nType == 2)
                    {
                        workBook = ExcelHelper.ToExcelPro(dt, "肥城市妇幼保健院诊疗项目（项目科室)表");
                    }
                    else if (nType == 3)
                    {
                        workBook = ExcelHelper.ToExcelPro(dt, "肥城市妇幼保健院诊疗项目（医师项目)表");
                    }
                    else if (nType == 4)
                    {
                        workBook = ExcelHelper.ToExcelPro(dt, "肥城市妇幼保健院诊疗项目（项目医师)表");
                    }

                    this.toolStripStatusLabel1.Text = "整理Excel表头，请等待......";
                    Application.DoEvents();

                    ISheet sheet1 = workBook.GetSheetAt(0);

                    // 
                    IRow row;
                    ICell cell;
                    int f = (Department.Value == "" ? 0 : 1);

                   
                    row = sheet1.GetRow(1);
                    cell = row.GetCell(0);
                    cell.SetCellValue("科室：" + Department.Text + " 日期范围：" + this.dateTimePicker1.Value.ToString("yyyy-MM-dd") + " ~ " + this.dateTimePicker2.Value.ToString("yyyy-MM-dd"));
                 
                    row = sheet1.GetRow(2);
                    row.Cells[0].SetCellValue("编码"); 
                    row.Cells[0].SetCellValue("名称");

                    if (nType == 1)
                    {
                        row.Cells[0].SetCellValue("科室编码");
                        row.Cells[1].SetCellValue("科室名称");
                    }
                    else if (nType == 2)
                    {
                        row.Cells[0].SetCellValue("项目编码");
                        row.Cells[1].SetCellValue("项目名称");
                    }
                    else if (nType == 3)
                    {
                        row.Cells[0].SetCellValue("医师编码");
                        row.Cells[1].SetCellValue("医师名称");
                    }
                    else if (nType == 4)
                    {
                        row.Cells[0].SetCellValue("项目编码");
                        row.Cells[1].SetCellValue("项目名称");
                    }



                    foreach (var item in row.Cells)
                    {
                        string cellValue = item.StringCellValue;
                        for (int i = 0; i < dt2.Rows.Count; i++)
                        {
                            DataRow dr2 = dt2.Rows[i];
                            string ColumnName = "";
                            string HeaderText = "";
                            if (nType == 1)
                            {
                                ColumnName = "col_" + Convert.ToString(dr2["项目编码"]).Trim();
                                HeaderText = Convert.ToString(dr2["项目名称"]).Trim();
                            }
                            else if (nType == 2)
                            {
                                ColumnName = "col_" + Convert.ToString(dr2["科码"]).Trim();
                                HeaderText = Convert.ToString(dr2["科名"]).Trim();
                            }
                            else if (nType == 3)
                            {
                                ColumnName = "col_" + Convert.ToString(dr2["项目编码"]).Trim();
                                HeaderText = Convert.ToString(dr2["项目名称"]).Trim();
                            }
                            else if (nType == 4)
                            {
                                ColumnName = "col_" + Convert.ToString(dr2["医师编码"]).Trim();
                                HeaderText = Convert.ToString(dr2["医师名称"]).Trim();
                            }


                            if (cellValue == ColumnName)
                            {
                                item.SetCellValue(HeaderText);
                                break;
                            }
                        }

                    }

                    System.IO.Directory.CreateDirectory(Application.StartupPath + @"\Excel");

                    string excelFile = Application.StartupPath + @"\Excel\诊疗处置_科室项目_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xls";

                    if (nType == 1)
                    {
                        if (this.comboBox3.SelectedIndex == 0)
                        {
                            excelFile = Application.StartupPath + @"\Excel\诊疗处置_科室项目_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xls";
                        }
                        else
                        {
                            excelFile = Application.StartupPath + @"\Excel\诊疗开方_科室项目_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xls";
                        }

                    }
                    else if (nType == 2)
                    {
                        if (this.comboBox3.SelectedIndex == 0)
                        {
                            excelFile = Application.StartupPath + @"\Excel\诊疗处置_项目科室_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xls";
                        }
                        else
                        {
                            excelFile = Application.StartupPath + @"\Excel\诊疗开方_项目科室_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xls";
                        }
                    }
                    else if (nType == 3)
                    {
                        if (this.comboBox3.SelectedIndex == 0)
                        {
                            excelFile = Application.StartupPath + @"\Excel\诊疗处置_医师项目_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xls";
                        }
                        else
                        {
                            excelFile = Application.StartupPath + @"\Excel\诊疗开方_医师项目_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xls";
                        }
                    }
                    else if (nType == 4)
                    {
                        if (this.comboBox3.SelectedIndex == 0)
                        {
                            excelFile = Application.StartupPath + @"\Excel\诊疗处置_项目医师_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xls";
                        }
                        else
                        {
                            excelFile = Application.StartupPath + @"\Excel\诊疗开方_项目医师_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xls";
                        }
                    }
                    FileStream stream = File.OpenWrite(excelFile); ;
                    workBook.Write(stream);
                    stream.Close();
                    MessageBox.Show("文件位置：" + excelFile, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);




                    #endregion
                }
            }
            else
            {
                MessageBox.Show("远程数据库连接失败！！！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            #endregion


        }
        private void UpdateDefault(SqlConnection conn)
        {
            string a = this.dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string b = this.dateTimePicker2.Value.ToString("yyyy-MM-dd");
            string sql = "";
            int nRet = 0;
            sql = " Update H4_收款记录 Set 科码='',科名='' Where 日期>='" + a + "' And 日期<='" + b + "'";
            nRet = ExecuteNonQuery(conn, sql);
            sql = " Update H4_收款记录 Set  H4_收款记录.收入编码= H0_收费项目.收入编码 From H0_收费项目 Where H0_收费项目.项目编码=H4_收款记录.项目编码 And  H4_收款记录.日期>='" + a + "' And H4_收款记录.日期<='" + b + "'";
            nRet = ExecuteNonQuery(conn, sql);
            sql = " Update H4_收款记录 Set  H4_收款记录.科码= H0_收费项目.科码, H4_收款记录.科名= H0_收费项目.科名  From H0_收费项目 Where H0_收费项目.项目编码=H4_收款记录.项目编码 And H4_收款记录.科码='' And  H4_收款记录.日期>='" + a + "' And H4_收款记录.日期<='" + b + "'";
            nRet = ExecuteNonQuery(conn, sql);
            sql = " Update H4_收款记录 Set  科码= 部门编码,  科名= 部门名称 Where 科码='' And  日期>='" + a + "' And 日期<='" + b + "'";
            nRet = ExecuteNonQuery(conn, sql);
        }



        private int ExecuteNonQuery(SqlConnection conn, string CommandText)
        {

            SqlCommand command = new SqlCommand();
            command.Connection = conn;
            command.CommandType = CommandType.Text;
            command.CommandText = CommandText;
            return command.ExecuteNonQuery();
        }

        private DataTable ExecuteDataTable(SqlConnection conn, string CommandText)
        {
            DataTable dt = new DataTable();
            SqlCommand command = new SqlCommand();
            command.Connection = conn;
            command.CommandType = CommandType.Text;
            command.CommandText = CommandText;
            SqlDataAdapter adapter = new SqlDataAdapter();
            adapter.SelectCommand = command;
            adapter.Fill(dt);
            return dt;
        }

    }
    public class SQLServer
    {
        public string SQL_Name { set; get; }
        public string SQL_ID { set; get; }
        public string SQL_PassWord { set; get; }
        public string SQL_DataBase { set; get; }
        public string ConnectionString { set; get; }
        public SqlConnection Connection { set; get; }
        public SqlTransaction Transaction { set; get; }
        public bool SelfConn { set; get; }
        public void TestConnection()
        {
            Connection = new SqlConnection();
            SelfConn = false;
            if (SQL_Name != "")
            {
                ConnectionString = string.Format("" +
                                   "Server = {0}; " +
                                   "Initial Catalog = {1}; " +
                                   "User ID = {2}; " +
                                   "Password = {3}; " +
                                   "max pool size = 800; min pool size = 300; Connect Timeout = 10",
                                   SQL_Name, SQL_DataBase, SQL_ID, SQL_PassWord);
                Connection.ConnectionString = ConnectionString;
                try
                {
                    Connection.Open();
                    SelfConn = true;
                    Connection.Close();
                }
                catch
                {
                    Connection = null;
                    SelfConn = false;
                }
            }

        }
    }
    public class ComboBoxItem
    {
        private string _text = "";
        private string _value = "";
        public string Text { get { return this._text; } set { this._text = value; } }
        public string Value { get { return this._value; } set { this._value = value; } }
        public override string ToString()
        {
            return this._text;
        }
    }
    public class s_m_total
    {
        public string 库码 { set; get; }
        public string 库名 { set; get; }
        public decimal 库码数量 { set; get; }
        public decimal 库码金额 { set; get; }
        public string 类码 { set; get; }
        public string 类名 { set; get; }
        public decimal 类码数量 { set; get; }
        public decimal 类码金额 { set; get; }
        public string 药品编码 { set; get; }
        public string 药品名称 { set; get; }
        public string 规格 { set; get; }
        public string 单位 { set; get; }
        public decimal 单价 { set; get; }
        public decimal 药品数量 { set; get; }
        public decimal 药品金额 { set; get; }
        public string 部门编码 { set; get; }
        public string 部门名称 { set; get; }
        public decimal 部门数量 { set; get; }
        public decimal 部门金额 { set; get; }
        public string 医师编码 { set; get; }
        public string 医师名称 { set; get; }
        public decimal 医师数量 { set; get; }
        public decimal 医师金额 { set; get; }
    }
    public class s_n_total
    {
        public string 科码 { set; get; }
        public string 科名 { set; get; }
        public decimal 科码数量 { set; get; }
        public decimal 科码金额 { set; get; }
        public string 项目编码 { set; get; }
        public string 项目名称 { set; get; }
        public decimal 项目数量 { set; get; }
        public decimal 项目金额 { set; get; }
        public string 部门编码 { set; get; }
        public string 部门名称 { set; get; }
        public decimal 部门数量 { set; get; }
        public decimal 部门金额 { set; get; }
        public string 医师编码 { set; get; }
        public string 医师名称 { set; get; }
        public decimal 医师数量 { set; get; }
        public decimal 医师金额 { set; get; }
    }



}
