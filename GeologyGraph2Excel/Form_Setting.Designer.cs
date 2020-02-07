namespace GeologyGraph2Excel
{
    partial class Form_Setting
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form_Setting));
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.rBA4 = new System.Windows.Forms.RadioButton();
            this.rBA3 = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button3 = new System.Windows.Forms.Button();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.textBox7 = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.chk_FCHD_Matck = new System.Windows.Forms.CheckBox();
            this.textBox8 = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.chk_DCBH = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(143, 80);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(83, 21);
            this.textBox4.TabIndex = 18;
            this.textBox4.Text = "0.5";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 83);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(131, 12);
            this.label4.TabIndex = 17;
            this.label4.Text = "标题与内容Y坐标容差：";
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(122, 23);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(104, 21);
            this.textBox3.TabIndex = 16;
            this.textBox3.Text = "钻孔编号";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 26);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(113, 12);
            this.label3.TabIndex = 15;
            this.label3.Text = "钻孔编号标题表头：";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(169, 47);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(51, 21);
            this.textBox2.TabIndex = 14;
            this.textBox2.Text = "297";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(128, 50);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 13;
            this.label2.Text = "高度：";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(59, 47);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(51, 21);
            this.textBox1.TabIndex = 12;
            this.textBox1.Text = "210";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 50);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 11;
            this.label1.Text = "宽度：";
            // 
            // rBA4
            // 
            this.rBA4.AutoSize = true;
            this.rBA4.Location = new System.Drawing.Point(19, 20);
            this.rBA4.Name = "rBA4";
            this.rBA4.Size = new System.Drawing.Size(65, 16);
            this.rBA4.TabIndex = 19;
            this.rBA4.TabStop = true;
            this.rBA4.Text = "A4 一页";
            this.rBA4.UseVisualStyleBackColor = true;
            this.rBA4.CheckedChanged += new System.EventHandler(this.rBA4_CheckedChanged);
            // 
            // rBA3
            // 
            this.rBA3.AutoSize = true;
            this.rBA3.Location = new System.Drawing.Point(105, 20);
            this.rBA3.Name = "rBA3";
            this.rBA3.Size = new System.Drawing.Size(65, 16);
            this.rBA3.TabIndex = 20;
            this.rBA3.TabStop = true;
            this.rBA3.Text = "A3 两页";
            this.rBA3.UseVisualStyleBackColor = true;
            this.rBA3.CheckedChanged += new System.EventHandler(this.rBA3_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button3);
            this.groupBox1.Controls.Add(this.rBA4);
            this.groupBox1.Controls.Add(this.rBA3);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(238, 84);
            this.groupBox1.TabIndex = 21;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "柱状图样式";
            // 
            // button3
            // 
            this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
            this.button3.Location = new System.Drawing.Point(181, 17);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(41, 24);
            this.button3.TabIndex = 21;
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // textBox5
            // 
            this.textBox5.Location = new System.Drawing.Point(122, 50);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(104, 21);
            this.textBox5.TabIndex = 23;
            this.textBox5.Text = "孔口高程";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 53);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(113, 12);
            this.label5.TabIndex = 22;
            this.label5.Text = "孔口高程标题表头：";
            // 
            // textBox6
            // 
            this.textBox6.Location = new System.Drawing.Point(122, 22);
            this.textBox6.Name = "textBox6";
            this.textBox6.Size = new System.Drawing.Size(90, 21);
            this.textBox6.TabIndex = 25;
            this.textBox6.Text = "地层编号";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(6, 25);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(113, 12);
            this.label6.TabIndex = 24;
            this.label6.Text = "地层编号标题表头：";
            // 
            // textBox7
            // 
            this.textBox7.Location = new System.Drawing.Point(122, 49);
            this.textBox7.Name = "textBox7";
            this.textBox7.Size = new System.Drawing.Size(104, 21);
            this.textBox7.TabIndex = 27;
            this.textBox7.Text = "分层厚度";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(6, 52);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(113, 12);
            this.label7.TabIndex = 26;
            this.label7.Text = "分层厚度标题表头：";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.textBox3);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.textBox4);
            this.groupBox2.Controls.Add(this.textBox5);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Location = new System.Drawing.Point(12, 102);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(238, 112);
            this.groupBox2.TabIndex = 22;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "标题";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.chk_FCHD_Matck);
            this.groupBox3.Controls.Add(this.textBox8);
            this.groupBox3.Controls.Add(this.label6);
            this.groupBox3.Controls.Add(this.textBox6);
            this.groupBox3.Controls.Add(this.label8);
            this.groupBox3.Controls.Add(this.textBox7);
            this.groupBox3.Controls.Add(this.label7);
            this.groupBox3.Location = new System.Drawing.Point(12, 225);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(238, 129);
            this.groupBox3.TabIndex = 22;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "内容";
            // 
            // chk_FCHD_Matck
            // 
            this.chk_FCHD_Matck.AutoSize = true;
            this.chk_FCHD_Matck.Location = new System.Drawing.Point(8, 75);
            this.chk_FCHD_Matck.Name = "chk_FCHD_Matck";
            this.chk_FCHD_Matck.Size = new System.Drawing.Size(216, 16);
            this.chk_FCHD_Matck.TabIndex = 29;
            this.chk_FCHD_Matck.Text = "分层厚度行数是否与其余列匹配一致";
            this.chk_FCHD_Matck.UseVisualStyleBackColor = true;
            this.chk_FCHD_Matck.CheckedChanged += new System.EventHandler(this.chk_FCHD_Matck_CheckedChanged);
            // 
            // textBox8
            // 
            this.textBox8.Location = new System.Drawing.Point(143, 98);
            this.textBox8.Name = "textBox8";
            this.textBox8.Size = new System.Drawing.Size(83, 21);
            this.textBox8.TabIndex = 25;
            this.textBox8.Text = "3.5";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(6, 101);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(131, 12);
            this.label8.TabIndex = 24;
            this.label8.Text = "表头与内容X坐标容差：";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 373);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(110, 39);
            this.button1.TabIndex = 23;
            this.button1.Text = "确 认";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(142, 373);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(108, 39);
            this.button2.TabIndex = 24;
            this.button2.Text = "取 消";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // chk_DCBH
            // 
            this.chk_DCBH.AutoSize = true;
            this.chk_DCBH.Location = new System.Drawing.Point(227, 250);
            this.chk_DCBH.Name = "chk_DCBH";
            this.chk_DCBH.Size = new System.Drawing.Size(15, 14);
            this.chk_DCBH.TabIndex = 28;
            this.chk_DCBH.UseVisualStyleBackColor = true;
            this.chk_DCBH.CheckedChanged += new System.EventHandler(this.chk_DCBH_CheckedChanged);
            // 
            // Form_Setting
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(264, 424);
            this.Controls.Add(this.chk_DCBH);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form_Setting";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "设置";
            this.Load += new System.EventHandler(this.Form_Setting_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton rBA4;
        private System.Windows.Forms.RadioButton rBA3;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox textBox5;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBox6;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textBox7;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.TextBox textBox8;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.CheckBox chk_DCBH;
        private System.Windows.Forms.CheckBox chk_FCHD_Matck;
        private System.Windows.Forms.Button button3;
    }
}