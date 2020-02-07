using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GeologyGraph2Excel
{
    public partial class Form_Setting : Form
    {
        public Form_Setting()
        {
            InitializeComponent();
        }

        private void Form_Setting_Load(object sender, EventArgs e)
        {
            textBox3.Text = GeologyGraph2Excel.BH_Keyword;
            textBox5.Text = GeologyGraph2Excel.KKGC_Keyword;
            textBox7.Text = GeologyGraph2Excel.FCHD_Keyword;
            textBox6.Text = GeologyGraph2Excel.DCBH_Keyword;
            textBox4.Text = GeologyGraph2Excel.Distance_alw.ToString();
            textBox8.Text = GeologyGraph2Excel.Distance_Column.ToString();
            chk_DCBH.Checked = GeologyGraph2Excel.isRead_DCBH;
            chk_FCHD_Matck.Checked = GeologyGraph2Excel.isMatch_FCHD;
            if (GeologyGraph2Excel.isA3 == true)
            {
                rBA3.Checked = true;
            }
            else
            {
                rBA4.Checked = true;
            }
            chk_DCBH_CheckedChanged(this,new EventArgs());
            textBox1.Text = GeologyGraph2Excel.PaperWidth.ToString();
            textBox2.Text = GeologyGraph2Excel.PaperHeight.ToString();
        }

        private void rBA4_CheckedChanged(object sender, EventArgs e)
        {
            textBox1.Text = "210";
            textBox2.Text = "297";
        }

        private void rBA3_CheckedChanged(object sender, EventArgs e)
        {
            textBox1.Text = "420";
            textBox2.Text = "297";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GeologyGraph2Excel.BH_Keyword = textBox3.Text;
            GeologyGraph2Excel.KKGC_Keyword = textBox5.Text;
            GeologyGraph2Excel.FCHD_Keyword = textBox7.Text;
            GeologyGraph2Excel.DCBH_Keyword = textBox6.Text;
            GeologyGraph2Excel.PaperWidth = Convert.ToDouble(textBox1.Text);
            GeologyGraph2Excel.PaperHeight = Convert.ToDouble(textBox2.Text);
            GeologyGraph2Excel.Distance_alw = Convert.ToDouble(textBox4.Text);
            GeologyGraph2Excel.Distance_Column = Convert.ToDouble(textBox8.Text);
            GeologyGraph2Excel.isRead_DCBH = chk_DCBH.Checked;
            GeologyGraph2Excel.isMatch_FCHD = chk_FCHD_Matck.Checked;
            if (rBA3.Checked == true)
            {
                GeologyGraph2Excel.isA3 = true;
            }
            else
            {
                GeologyGraph2Excel.isA3 = false;
            }
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void chk_DCBH_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_DCBH.Checked == true)
            {
                textBox6.Enabled = true;
            }
            else
            {
                textBox6.Enabled = false;
            }
        }

        private void chk_FCHD_Matck_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_FCHD_Matck.Checked == false)
            {
                MessageBox.Show("不强制匹配的情况下，能够提高识别的成功率，" +
                    "适用于某些柱状图因最后几行空间不足的情况下，只标注了地层编号和岩土名称，" +
                    "却把分层厚度放在下一页才标注，但有较大可能性读取的图表会有误," +
                    "并且后期合并列表的时候很大可能会出错，请慎重不勾选！");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            List<double> a= GeologyGraph2Excel.Measure();
            if (a != null)
            {
                textBox1.Text = a[0].ToString();
                textBox2.Text = a[1].ToString();
            }
            this.Show();
        }
    }
}
