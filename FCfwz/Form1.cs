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

namespace FCfwz
{
    public partial class Form1 : Form
    {
        public SQLServer MySQLServer = new SQLServer();
        public string XMLPath = "";
        public Form1()
        {
            InitializeComponent();
        }


        private void Form1_Load(object sender, EventArgs e)
        {

            string ret = GetAppSettings();
            if (ret  == "")
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
                MessageBox.Show(ret,"错误",MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

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
            MySQLServer.GetConnection();
            if (MySQLServer.SelfConn == true) { 
                MessageBox.Show("远程数据库连接成功！！！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            } else {
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
            else {
                MessageBox.Show("配置信息保存成功！！！","提示",MessageBoxButtons.OK,MessageBoxIcon.Asterisk);
            }
        }
        //查询
        private void button4_Click(object sender, EventArgs e)
        {

        }
        //导出
        private void button5_Click(object sender, EventArgs e)
        {

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
        public void GetConnection()
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
                }
                catch
                {
                    Connection = null;
                    SelfConn = false;
                }
            }

        }

    }
}
