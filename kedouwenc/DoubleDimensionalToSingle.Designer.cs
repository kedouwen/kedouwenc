namespace kedouwenc
{
    partial class DoubleDimensionalToSingle
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DoubleDimensionalToSingle));
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "选择数据区域";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(96, 24);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(118, 21);
            this.textBox1.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(212, 23);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(17, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "-";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Location = new System.Drawing.Point(13, 89);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(216, 292);
            this.checkedListBox1.TabIndex = 3;
            this.checkedListBox1.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.checkedListBox1_ItemCheck);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 70);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(113, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "请选择待转换的字段";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(20, 407);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(209, 12);
            this.label3.TabIndex = 5;
            this.label3.Text = "提示：右图中苹果桔子梨为待转换字段";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(22, 438);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 6;
            this.button2.Text = "确定";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(133, 437);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 7;
            this.button3.Text = "取消";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // label4
            // 
            this.label4.Image = ((System.Drawing.Image)(resources.GetObject("label4.Image")));
            this.label4.Location = new System.Drawing.Point(237, 23);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(238, 133);
            this.label4.TabIndex = 8;
            this.label4.Text = "示例";
            this.label4.UseCompatibleTextRendering = true;
            // 
            // label6
            // 
            this.label6.Image = ((System.Drawing.Image)(resources.GetObject("label6.Image")));
            this.label6.Location = new System.Drawing.Point(252, 193);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(198, 299);
            this.label6.TabIndex = 10;
            // 
            // label5
            // 
            this.label5.AllowDrop = true;
            this.label5.Image = ((System.Drawing.Image)(resources.GetObject("label5.Image")));
            this.label5.Location = new System.Drawing.Point(336, 148);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(48, 44);
            this.label5.TabIndex = 9;
            // 
            // DoubleDimensionalToSingle
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(487, 510);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.checkedListBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label1);
            this.Name = "DoubleDimensionalToSingle";
            this.Text = "二维表转一维表";
            this.Load += new System.EventHandler(this.DoubleDimensionalToSingle_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        //加入事件 
        void checkedListBox1_ItemCheck(object sender, System.Windows.Forms.ItemCheckEventArgs e)
        {
            if (e.Index == 0 & e.CurrentValue == 0)
            {
                e.NewValue = 0;
            }    
            //throw new System.NotImplementedException();
        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
    }
}