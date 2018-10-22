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
using System.Threading.Tasks;
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
                MessageBox.Show(ret, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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


        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            this.splitContainer1.SplitterDistance = 300;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (this.splitContainer1.SplitterDistance == 300)
            {
                this.splitContainer1.SplitterDistance = 30;
                this.button1.Text = ">>";
                this.groupBox1.Visible = false;
                this.groupBox2.Visible = false;
            }
            else
            {
                this.splitContainer1.SplitterDistance = 300;
                this.button1.Text = "配置项设置，单击可以隐藏";
                this.groupBox1.Visible = true;
                this.groupBox2.Visible = true;
            }
        }

        private void toolTip1_Popup(object sender, PopupEventArgs e)
        {

        }

        private void button1_MouseEnter(object sender, EventArgs e)
        {
            if (this.splitContainer1.SplitterDistance == 30)
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
        //查询
        private void button4_Click(object sender, EventArgs e)
        {
            List<s_m_total> m = new List<s_m_total>();
            List<s_n_total> n = new List<s_n_total>();

            if (this.radioButton1.Checked == true)
            {
                m = Get_m_total();
                this.dataGridView1.DataSource = null;
                this.dataGridView1.DataSource = m;
            }

            if (this.radioButton2.Checked == true)
            {
                n = Get_n_total();
                this.dataGridView1.DataSource = null;
                this.dataGridView1.DataSource = n;
            }
        }
        //导出
        private void button5_Click(object sender, EventArgs e)
        {
            List<s_m_total> m = new List<s_m_total>();
            List<s_n_total> n = new List<s_n_total>();

            if (this.radioButton1.Checked == true)
            {
                m = Get_m_total();
                Excel_m_total(m);
            }

            if (this.radioButton2.Checked == true)
            {
                n = Get_n_total();
                Excel_n_total(n);
            }
        }

        private void Excel_m_total(List<s_m_total> m)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("类别", System.Type.GetType("System.String"));
            dt.Columns.Add("药品", System.Type.GetType("System.String"));
            dt.Columns.Add("规格", System.Type.GetType("System.String"));
            dt.Columns.Add("单位", System.Type.GetType("System.String"));
            dt.Columns.Add("单价", System.Type.GetType("System.Decimal"));
            dt.Columns.Add("药品数量", System.Type.GetType("System.Int32"));
            dt.Columns.Add("药品金额", System.Type.GetType("System.Decimal"));
            dt.Columns.Add("部门", System.Type.GetType("System.String"));
            dt.Columns.Add("部门数量", System.Type.GetType("System.Int32"));
            dt.Columns.Add("部门金额", System.Type.GetType("System.Decimal"));
            dt.Columns.Add("医师", System.Type.GetType("System.String"));
            dt.Columns.Add("医师数量", System.Type.GetType("System.Int32"));
            dt.Columns.Add("医师金额", System.Type.GetType("System.Decimal"));
            foreach (s_m_total item in m)
            {
                dt.Rows.Add(item.类名, item.药品名称, item.规格, item.单位, item.单价,
                            Convert.ToInt32(item.数量), item.金额,
                            item.部门名称, Convert.ToInt32(item.部门数量), item.部门金额,
                            item.医师名称, Convert.ToInt32(item.医师数量), item.医师金额);
            }
            IWorkbook workBook = new HSSFWorkbook();
            workBook = ExcelHelper.ToExcel(dt, "药品分类汇总");

            ISheet sheet1 = workBook.GetSheetAt(0);
            sheet1.SetColumnWidth(0, 20 * 256);     //类名
            sheet1.SetColumnWidth(1, 30 * 256);     //药品名称
            sheet1.SetColumnWidth(2, 15 * 256);     //规格
            sheet1.SetColumnWidth(3, 6 * 256);      //单位
            sheet1.SetColumnWidth(4, 12 * 256);     //单价
            sheet1.SetColumnWidth(5, 12 * 256);     //数量
            sheet1.SetColumnWidth(6, 12 * 256);     //金额
            sheet1.SetColumnWidth(7, 20 * 256);     //部门
            sheet1.SetColumnWidth(8, 12 * 256);     //数量
            sheet1.SetColumnWidth(9, 12 * 256);     //金额
            sheet1.SetColumnWidth(10, 12 * 256);    //医师
            sheet1.SetColumnWidth(11, 12 * 256);    //数量
            sheet1.SetColumnWidth(12, 12 * 256);    //金额

            //整理表头 
            IRow row;
            ICell cell;
            ComboBoxItem Department = (ComboBoxItem)this.comboBox1.SelectedItem;
            #region 第一行
            row = sheet1.GetRow(1);
            cell = row.GetCell(0);
            cell.SetCellValue("科室：" + Department.Text + " 日期范围：" + this.dateTimePicker1.Value.ToString("yyyy-MM-dd") + " ~ " + this.dateTimePicker2.Value.ToString("yyyy-MM-dd"));
            #endregion
            #region 第二行
            row = sheet1.GetRow(2);
            cell = row.GetCell(0);
            cell.SetCellValue("药品");
            cell = row.GetCell(1);
            cell.SetCellValue("药品");
            cell = row.GetCell(2);
            cell.SetCellValue("药品");
            cell = row.GetCell(3);
            cell.SetCellValue("药品");
            cell = row.GetCell(4);
            cell.SetCellValue("药品");
            cell = row.GetCell(5);
            cell.SetCellValue("药品");
            cell = row.GetCell(6);
            cell.SetCellValue("药品");

            cell = row.GetCell(7);
            cell.SetCellValue("部门");
            cell = row.GetCell(8);
            cell.SetCellValue("部门");
            cell = row.GetCell(9);
            cell.SetCellValue("部门");

            cell = row.GetCell(10);
            cell.SetCellValue("医师");
            cell = row.GetCell(11);
            cell.SetCellValue("医师");
            cell = row.GetCell(12);
            cell.SetCellValue("医师");
            #endregion 
            #region 第三行
            row = sheet1.GetRow(3);
            cell = row.GetCell(0);
            cell.SetCellValue("类别");
            cell = row.GetCell(1);
            cell.SetCellValue("药名");
            cell = row.GetCell(2);
            cell.SetCellValue("规格");
            cell = row.GetCell(3);
            cell.SetCellValue("单位");
            cell = row.GetCell(4);
            cell.SetCellValue("单价");
            cell = row.GetCell(5);
            cell.SetCellValue("数量");
            cell = row.GetCell(6);
            cell.SetCellValue("金额");

            cell = row.GetCell(7);
            cell.SetCellValue("名称");
            cell = row.GetCell(8);
            cell.SetCellValue("数量");
            cell = row.GetCell(9);
            cell.SetCellValue("金额");

            cell = row.GetCell(10);
            cell.SetCellValue("姓名");
            cell = row.GetCell(11);
            cell.SetCellValue("数量");
            cell = row.GetCell(12);
            cell.SetCellValue("金额");
            #endregion 
            #region 合并表头第一行
            sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 0, 6));
            sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 7, 9));
            sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 10, 12));
            #endregion
            #region 正文合并

            int start = 0;      //记录同组开始行号
            int end = 0;        //记录同组结束行号
            string temp = "";
            for (int j = 0; j < 12; j++)
            {
                start = 4;  //记录同组开始行号
                end = 4;    //记录同组结束行号

                for (int i = 0; i < m.Count  ; i++)
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
            System.IO.Directory.CreateDirectory(Application.StartupPath + "//Excel");
            string excelFile = Application.StartupPath + "//Excel//药品_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xls";
            FileStream stream = File.OpenWrite(excelFile); ;
            workBook.Write(stream);
            stream.Close();
            MessageBox.Show("文件位置：" + excelFile, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Excel_n_total(List<s_n_total> n)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("科室", System.Type.GetType("System.String"));
            dt.Columns.Add("项目", System.Type.GetType("System.String"));
            dt.Columns.Add("项目数量", System.Type.GetType("System.Int32"));
            dt.Columns.Add("项目金额", System.Type.GetType("System.Decimal"));
            dt.Columns.Add("部门", System.Type.GetType("System.String"));
            dt.Columns.Add("部门数量", System.Type.GetType("System.Int32"));
            dt.Columns.Add("部门金额", System.Type.GetType("System.Decimal"));
            dt.Columns.Add("医师", System.Type.GetType("System.String"));
            dt.Columns.Add("医师数量", System.Type.GetType("System.Int32"));
            dt.Columns.Add("医师金额", System.Type.GetType("System.Decimal"));
            foreach (s_n_total item in n)
            {
                dt.Rows.Add(item.科名, item.项目名称, Convert.ToInt32(item.数量), item.金额,
                            item.部门名称, Convert.ToInt32(item.部门数量), item.部门金额,
                            item.医师名称, Convert.ToInt32(item.医师数量), item.医师金额);
            }
            IWorkbook workBook = new HSSFWorkbook();
            workBook = ExcelHelper.ToExcel(dt, "诊疗分类汇总");

            ISheet sheet1 = workBook.GetSheetAt(0);
            sheet1.SetColumnWidth(0, 20 * 256);     //科名
            sheet1.SetColumnWidth(1, 50 * 256);     //项目名称
            sheet1.SetColumnWidth(2, 12 * 256);     //数量
            sheet1.SetColumnWidth(3, 12 * 256);     //金额
            sheet1.SetColumnWidth(4, 20 * 256);     //部门
            sheet1.SetColumnWidth(5, 12 * 256);     //数量
            sheet1.SetColumnWidth(6, 12 * 256);     //金额
            sheet1.SetColumnWidth(7, 12 * 256);     //医师
            sheet1.SetColumnWidth(8, 12 * 256);     //数量
            sheet1.SetColumnWidth(9, 12 * 256);     //金额


            //整理表头 
            IRow row;
            ICell cell;
            ComboBoxItem Department = (ComboBoxItem)this.comboBox1.SelectedItem;
            #region 第一行
            row = sheet1.GetRow(1);
            cell = row.GetCell(0);
            cell.SetCellValue("科室：" + Department.Text + " 日期范围：" + this.dateTimePicker1.Value.ToString("yyyy-MM-dd") + " ~ " + this.dateTimePicker2.Value.ToString("yyyy-MM-dd"));
            #endregion
            #region 第二行
            row = sheet1.GetRow(2);
            cell = row.GetCell(0);
            cell.SetCellValue("科名");
            cell = row.GetCell(1);
            cell.SetCellValue("项目");
            cell = row.GetCell(2);
            cell.SetCellValue("项目");
            cell = row.GetCell(3);
            cell.SetCellValue("项目");

            cell = row.GetCell(4);
            cell.SetCellValue("部门");
            cell = row.GetCell(5);
            cell.SetCellValue("部门");
            cell = row.GetCell(6);
            cell.SetCellValue("部门");

            cell = row.GetCell(7);
            cell.SetCellValue("医师");
            cell = row.GetCell(8);
            cell.SetCellValue("医师");
            cell = row.GetCell(9);
            cell.SetCellValue("医师");
            #endregion 
            #region 第三行
            row = sheet1.GetRow(3);
            cell = row.GetCell(0);
            cell.SetCellValue("科名");

            cell = row.GetCell(1);
            cell.SetCellValue("名称");
            cell = row.GetCell(2);
            cell.SetCellValue("数量");
            cell = row.GetCell(3);
            cell.SetCellValue("金额");

            cell = row.GetCell(4);
            cell.SetCellValue("名称");
            cell = row.GetCell(5);
            cell.SetCellValue("数量");
            cell = row.GetCell(6);
            cell.SetCellValue("金额");

            cell = row.GetCell(7);
            cell.SetCellValue("姓名");
            cell = row.GetCell(8);
            cell.SetCellValue("数量");
            cell = row.GetCell(9);
            cell.SetCellValue("金额");
            #endregion 
            #region 合并表头第一行
            sheet1.AddMergedRegion(new CellRangeAddress(2, 3, 0, 0));
            sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 1, 3));
            sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 4, 6));
            sheet1.AddMergedRegion(new CellRangeAddress(2, 2, 7, 9));
            #endregion
            #region 正文合并

            int start = 0;      //记录同组开始行号
            int end = 0;        //记录同组结束行号
            string temp = "";
            for (int j = 0; j < 8; j++)
            {
                start = 4;  //记录同组开始行号
                end = 4;    //记录同组结束行号

                for (int i = 0; i < n.Count  ; i++)
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
            System.IO.Directory.CreateDirectory(Application.StartupPath + "//Excel");
            string excelFile = Application.StartupPath + "//Excel//诊疗_" + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".xls";
            FileStream stream = File.OpenWrite(excelFile); ;
            workBook.Write(stream);
            stream.Close();
            MessageBox.Show("文件位置：" + excelFile, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private string GetAppSettings()
        {
            string Ret = "";
            try
            {
                Configuration config = System.Configuration.ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

                if (config.AppSettings.Settings["SQL_Name"] != null)
                {
                    MySQLServer.SQL_Name = config.AppSettings.Settings["SQL_Name"].Value;
                }
                if (config.AppSettings.Settings["SQL_ID"] != null)
                {
                    MySQLServer.SQL_ID = config.AppSettings.Settings["SQL_ID"].Value;
                }
                if (config.AppSettings.Settings["SQL_PassWord"] != null)
                {
                    MySQLServer.SQL_PassWord = config.AppSettings.Settings["SQL_PassWord"].Value;
                }
                if (config.AppSettings.Settings["SQL_DataBase"] != null)
                {
                    MySQLServer.SQL_DataBase = config.AppSettings.Settings["SQL_DataBase"].Value;
                }
            }
            catch (Exception ex)
            {

                Ret = ex.Message.ToString();
            }
            return Ret;
        }

        private string SetAppSettings()
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

        public static SqlConnection GetSqlConnection(string SQL_Name, string SQL_DataBase, string SQL_ID, string SQL_PassWord, int Timeout = 20)
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

        private List<ComboBoxItem> GetAllDepartment()
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
            else
            {
                MessageBox.Show("远程数据库连接失败！！！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return Department;
        }

        private List<s_m_total> Get_m_total()
        {
            ComboBoxItem Department = (ComboBoxItem)this.comboBox1.SelectedItem;

            List<s_m_total> m_total = new List<s_m_total>();
            if (MySQLServer.SelfConn == false)
            {
                MySQLServer.TestConnection();
            }
            if (MySQLServer.SelfConn == true)
            {
                using (var conn = GetSqlConnection(MySQLServer.SQL_Name, MySQLServer.SQL_DataBase, MySQLServer.SQL_ID, MySQLServer.SQL_PassWord, 1000))
                {
                    string t = "1";
                    string a = this.dateTimePicker1.Value.ToString("yyyy-MM-dd");
                    string b = this.dateTimePicker2.Value.ToString("yyyy-MM-dd");
                    string c = Department.Value;
                    DataTable dt = new DataTable();
                    try
                    {
                        SqlCommand command = new SqlCommand();
                        command.Connection = conn;
                        command.CommandType = CommandType.StoredProcedure;
                        command.CommandText = "s_m_total";
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

                    foreach (DataRow dr in dt.Rows)
                    {
                        s_m_total model = new s_m_total();
                        if (dr["类码"] != DBNull.Value) model.类码 = Convert.ToString(dr["类码"]).Trim();
                        if (dr["类名"] != DBNull.Value) model.类名 = Convert.ToString(dr["类名"]).Trim();
                        if (dr["药品编码"] != DBNull.Value) model.药品编码 = Convert.ToString(dr["药品编码"]).Trim();
                        if (dr["药品名称"] != DBNull.Value) model.药品名称 = Convert.ToString(dr["药品名称"]).Trim();
                        if (dr["规格"] != DBNull.Value) model.规格 = Convert.ToString(dr["规格"]).Trim();
                        if (dr["单位"] != DBNull.Value) model.单位 = Convert.ToString(dr["单位"]).Trim();
                        if (dr["单价"] != DBNull.Value) model.单价 = Convert.ToDecimal(dr["单价"]);
                        if (dr["sl_t3"] != DBNull.Value) model.数量 = Convert.ToDecimal(dr["sl_t3"]);
                        if (dr["je_t3"] != DBNull.Value) model.金额 = Convert.ToDecimal(dr["je_t3"]);
                        if (dr["部门编码"] != DBNull.Value) model.部门编码 = Convert.ToString(dr["部门编码"]).Trim();
                        if (dr["部门名称"] != DBNull.Value) model.部门名称 = Convert.ToString(dr["部门名称"]).Trim();
                        if (dr["sl_t4"] != DBNull.Value) model.部门数量 = Convert.ToDecimal(dr["sl_t4"]);
                        if (dr["je_t4"] != DBNull.Value) model.部门金额 = Convert.ToDecimal(dr["je_t4"]);
                        if (dr["医师编码"] != DBNull.Value) model.医师编码 = Convert.ToString(dr["部门编码"]).Trim();
                        if (dr["医师名称"] != DBNull.Value) model.医师名称 = Convert.ToString(dr["部门名称"]).Trim();
                        if (dr["sl_t5"] != DBNull.Value) model.医师数量 = Convert.ToDecimal(dr["sl_t5"]);
                        if (dr["je_t5"] != DBNull.Value) model.医师金额 = Convert.ToDecimal(dr["je_t5"]);
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

        private List<s_n_total> Get_n_total()
        {

            ComboBoxItem Department = (ComboBoxItem)this.comboBox1.SelectedItem;


            List<s_n_total> n_total = new List<s_n_total>();
            if (MySQLServer.SelfConn == false)
            {
                MySQLServer.TestConnection();
            }
            if (MySQLServer.SelfConn == true)
            {
                using (var conn = GetSqlConnection(MySQLServer.SQL_Name, MySQLServer.SQL_DataBase, MySQLServer.SQL_ID, MySQLServer.SQL_PassWord, 1000))
                {
                    string t = "1";
                    string a = this.dateTimePicker1.Value.ToString("yyyy-MM-dd");
                    string b = this.dateTimePicker2.Value.ToString("yyyy-MM-dd");
                    string c = Department.Value;
                    DataTable dt = new DataTable();
                    try
                    {
                        SqlCommand command = new SqlCommand();
                        command.Connection = conn;
                        command.CommandType = CommandType.StoredProcedure;
                        command.CommandText = "s_n_total";
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
                    foreach (DataRow dr in dt.Rows)
                    {
                        s_n_total model = new s_n_total();
                        if (dr["科码"] != DBNull.Value) model.科码 = Convert.ToString(dr["科码"]).Trim();
                        if (dr["科名"] != DBNull.Value) model.科名 = Convert.ToString(dr["科名"]).Trim();
                        if (dr["项目编码"] != DBNull.Value) model.项目编码 = Convert.ToString(dr["项目编码"]).Trim();
                        if (dr["项目名称"] != DBNull.Value) model.项目名称 = Convert.ToString(dr["项目名称"]).Trim();
                        if (dr["sl_t3"] != DBNull.Value) model.数量 = Convert.ToDecimal(dr["sl_t3"]);
                        if (dr["je_t3"] != DBNull.Value) model.金额 = Convert.ToDecimal(dr["je_t3"]);
                        if (dr["部门编码"] != DBNull.Value) model.部门编码 = Convert.ToString(dr["部门编码"]).Trim();
                        if (dr["部门名称"] != DBNull.Value) model.部门名称 = Convert.ToString(dr["部门名称"]).Trim();
                        if (dr["sl_t4"] != DBNull.Value) model.部门数量 = Convert.ToDecimal(dr["sl_t4"]);
                        if (dr["je_t4"] != DBNull.Value) model.部门金额 = Convert.ToDecimal(dr["je_t4"]);
                        if (dr["医师编码"] != DBNull.Value) model.医师编码 = Convert.ToString(dr["部门编码"]).Trim();
                        if (dr["医师名称"] != DBNull.Value) model.医师名称 = Convert.ToString(dr["部门名称"]).Trim();
                        if (dr["sl_t5"] != DBNull.Value) model.医师数量 = Convert.ToDecimal(dr["sl_t5"]);
                        if (dr["je_t5"] != DBNull.Value) model.医师金额 = Convert.ToDecimal(dr["je_t5"]);
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

        private void button6_Click(object sender, EventArgs e)
        {
            this.comboBox1.Items.Clear();
            List<ComboBoxItem> Departments = GetAllDepartment();
            foreach (ComboBoxItem Department in Departments)
            {
                this.comboBox1.Items.Add(Department);
            }
            if (this.comboBox1.Items.Count > 0) { this.comboBox1.SelectedIndex = 0; }
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
                                   "max pool size = 800; min pool size = 300; Connect Timeout = 20",
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
        public string 类码 { set; get; }
        public string 类名 { set; get; }
        public string 药品编码 { set; get; }
        public string 药品名称 { set; get; }
        public string 规格 { set; get; }
        public string 单位 { set; get; }
        public decimal 单价 { set; get; }
        public decimal 数量 { set; get; }
        public decimal 金额 { set; get; }
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
        public string 项目编码 { set; get; }
        public string 项目名称 { set; get; }
        public decimal 数量 { set; get; }
        public decimal 金额 { set; get; }
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
