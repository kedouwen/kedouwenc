namespace kedouwenc
{
    partial class TextToPinyinForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TextToPinyinForm));
            this.labelInput = new System.Windows.Forms.Label();
            this.buttonCnToPinyin = new System.Windows.Forms.Button();
            this.buttonTestDict = new System.Windows.Forms.Button();
            this.buttonSegm = new System.Windows.Forms.Button();
            this.checkBoxWithTone = new System.Windows.Forms.CheckBox();
            this.buttonSelectDict = new System.Windows.Forms.Button();
            this.openFileDialogDict = new System.Windows.Forms.OpenFileDialog();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.button4 = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.button5 = new System.Windows.Forms.Button();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // labelInput
            // 
            this.labelInput.AutoSize = true;
            this.labelInput.Location = new System.Drawing.Point(12, 9);
            this.labelInput.Name = "labelInput";
            this.labelInput.Size = new System.Drawing.Size(113, 12);
            this.labelInput.TabIndex = 2;
            this.labelInput.Text = "要处理的单元格区域";
            // 
            // buttonCnToPinyin
            // 
            this.buttonCnToPinyin.Location = new System.Drawing.Point(12, 35);
            this.buttonCnToPinyin.Name = "buttonCnToPinyin";
            this.buttonCnToPinyin.Size = new System.Drawing.Size(99, 23);
            this.buttonCnToPinyin.TabIndex = 4;
            this.buttonCnToPinyin.Text = "转换为拼音";
            this.buttonCnToPinyin.UseVisualStyleBackColor = true;
            this.buttonCnToPinyin.Click += new System.EventHandler(this.buttonCnToPinyin_Click);
            // 
            // buttonTestDict
            // 
            this.buttonTestDict.Location = new System.Drawing.Point(17, 54);
            this.buttonTestDict.Name = "buttonTestDict";
            this.buttonTestDict.Size = new System.Drawing.Size(105, 23);
            this.buttonTestDict.TabIndex = 5;
            this.buttonTestDict.Text = "测试读取词典";
            this.buttonTestDict.UseVisualStyleBackColor = true;
            this.buttonTestDict.Click += new System.EventHandler(this.buttonTestDict_Click);
            // 
            // buttonSegm
            // 
            this.buttonSegm.Location = new System.Drawing.Point(30, 252);
            this.buttonSegm.Name = "buttonSegm";
            this.buttonSegm.Size = new System.Drawing.Size(105, 23);
            this.buttonSegm.TabIndex = 6;
            this.buttonSegm.Text = "正向分词";
            this.buttonSegm.UseVisualStyleBackColor = true;
            this.buttonSegm.Click += new System.EventHandler(this.buttonSegm_Click);
            // 
            // checkBoxWithTone
            // 
            this.checkBoxWithTone.AutoSize = true;
            this.checkBoxWithTone.Location = new System.Drawing.Point(59, 13);
            this.checkBoxWithTone.Name = "checkBoxWithTone";
            this.checkBoxWithTone.Size = new System.Drawing.Size(72, 16);
            this.checkBoxWithTone.TabIndex = 7;
            this.checkBoxWithTone.Text = "包含音调";
            this.checkBoxWithTone.UseVisualStyleBackColor = true;
            // 
            // buttonSelectDict
            // 
            this.buttonSelectDict.Location = new System.Drawing.Point(17, 20);
            this.buttonSelectDict.Name = "buttonSelectDict";
            this.buttonSelectDict.Size = new System.Drawing.Size(105, 23);
            this.buttonSelectDict.TabIndex = 9;
            this.buttonSelectDict.Text = "自选词典";
            this.buttonSelectDict.UseVisualStyleBackColor = true;
            this.buttonSelectDict.Click += new System.EventHandler(this.buttonSelectDict_Click);
            // 
            // openFileDialogDict
            // 
          //  this.openFileDialogDict.FileName = "openFileDialog1";
            
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(277, 43);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(21, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "-";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(25, 45);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(251, 21);
            this.textBox1.TabIndex = 10;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(30, 291);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(105, 23);
            this.button2.TabIndex = 1;
            this.button2.Text = "逆向分词";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(30, 333);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(105, 23);
            this.button3.TabIndex = 2;
            this.button3.Text = "双向分词";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Location = new System.Drawing.Point(14, 231);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(142, 136);
            this.groupBox1.TabIndex = 14;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "分词";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.buttonSelectDict);
            this.groupBox2.Controls.Add(this.buttonTestDict);
            this.groupBox2.Location = new System.Drawing.Point(174, 231);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(142, 136);
            this.groupBox2.TabIndex = 15;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "词典";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(11, 76);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(100, 23);
            this.button4.TabIndex = 16;
            this.button4.Text = "转换为首字母";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.button4);
            this.groupBox3.Controls.Add(this.checkBoxWithTone);
            this.groupBox3.Controls.Add(this.buttonCnToPinyin);
            this.groupBox3.Location = new System.Drawing.Point(13, 82);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(142, 115);
            this.groupBox3.TabIndex = 17;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "转换";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.button5);
            this.groupBox4.Location = new System.Drawing.Point(174, 83);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(147, 114);
            this.groupBox4.TabIndex = 18;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "统计";
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(23, 35);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(93, 23);
            this.button5.TabIndex = 0;
            this.button5.Text = "分词统计";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // TextToPinyinForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(328, 397);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.buttonSegm);
            this.Controls.Add(this.labelInput);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBox2);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "TextToPinyinForm";
            this.Text = "中文分词/转换拼音/转换首字母";
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelInput;
        private System.Windows.Forms.Button buttonCnToPinyin;
        private System.Windows.Forms.Button buttonTestDict;
        private System.Windows.Forms.Button buttonSegm;
        private System.Windows.Forms.CheckBox checkBoxWithTone;
        private System.Windows.Forms.Button buttonSelectDict;
        private System.Windows.Forms.OpenFileDialog openFileDialogDict;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button button5;
    }
}

