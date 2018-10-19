using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FCfwz
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
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
            else {
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
            if (this.splitContainer1.SplitterDistance == 30) { 
                this.toolTip1.IsBalloon = false;
                this.toolTip1.UseFading = true;
                this.toolTip1.Show("配置项设置，单击可以隐藏", this.button1); 
            } 
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            this.toolTip1.Hide(this.button1);     //隐藏提示窗口
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            this.splitContainer1.SplitterDistance = 300;
        }
    }

    public class SQLServer {
        public string SQL_Name { set; get; }
        public string SQL_ID { set; get; }
        public string SQL_PassWord { set; get; }
        public string SQL_DataBase { set; get; } 
    }
}
