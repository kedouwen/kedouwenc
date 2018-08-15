using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace kedouwenc
{
    public partial class TextToPinyinForm : Form
    {
        //词典的文件名，含路径
        string filename;
        Excel.Worksheet sht = Globals.ThisAddIn.Application.ActiveSheet;
        Excel.Range rng1;
        object[,] arr1; 


        public TextToPinyinForm()
        {
            InitializeComponent();
            //初始化词典文件路径；
           // filename = @"Dictionary\\Dictionary.xml";
            filename = @AppDomain.CurrentDomain.BaseDirectory + "PinyinService\\Dictionary\\Dictionary.xml";

          //   MessageBox.Show(filename);
        }
        



        //测试各个分词模块
        private string SegWords(string str, int seglb)
        {
            string SegStr = "";

            //生成词典
            Entity.PinyinDictionary dict = new Entity.PinyinDictionary(@filename);

            //只取词典中的中文词汇，无需拼音
            List<string> wordList = dict.Dictionary.Keys.ToList<string>();

            if (seglb == 1)
            {
                //进行正向分词
                List<string> wordsLeft = Helper.Segmentation.SegMMLeftToRight(str, ref wordList);

                //判断分词是否正常返回
                if (wordsLeft == null)
                {
                    return "正向分词模块执行失败";
                }

                //将正向分词结果显示出来
                SegStr += "正向分词：";
                foreach (string word in wordsLeft)
                {
                    SegStr += word;
                    SegStr += ",";
                }

                //换行
                //SegStr += "\r\n";
                //SegStr += "\\";
                return SegStr;
            }

            else if (seglb == 2)
            {
                //进行逆向分词
                List<string> wordsRight = Helper.Segmentation.SegMMRightToLeft(str, ref wordList);

                //判断分词是否正常返回
                if (wordsRight == null)
                {
                    return "逆向分词模块执行失败";
                }

                //将正向分词结果显示出来
                SegStr += "逆向分词：";
                foreach (string word in wordsRight)
                {
                    SegStr += word;
                    SegStr += ",";
                }
                //换行
                //SegStr += "\r\n";
                // SegStr += "\\";
                return SegStr;
            }

            else
            {
                //进行双向分词
                List<string> wordsDouble = Helper.Segmentation.SegMMDouble(str, ref wordList);

                //判断分词是否正常返回
                if (wordsDouble == null)
                {
                    return "双向分词模块执行失败";
                }

                //将正向分词结果显示出来
                SegStr += "双向分词：";
                foreach (string word in wordsDouble)
                {
                    SegStr += word;
                    SegStr += ",";
                }
                return SegStr;
            }

        }

        //测试读取词典
        private void ReadDict()
        {

            string dictext = "";

            //读取词典
            Entity.PinyinDictionary dict = new Entity.PinyinDictionary(@filename);

            //显示词典条数

            dictext = "词典读取成功，获得词条：";
            dictext += dict.Dictionary.Count;
            dictext += "条。";
            dictext += "\r\n";

            //提示只显示前100条，不然程序会卡死
            if (dict.Dictionary.Count > 50)
            {
                dictext += "只显示词典的前50条，太多程序会卡死：";
                dictext += "\r\n";
            }
            else
            {
                dictext += "词典所有内容如下：";
                dictext += "\r\n";
            }

            //将词典内容显示在窗口中，只显示前50条
            int i = 0;
            foreach (KeyValuePair<string, string> pair in dict.Dictionary)
            {
                dictext += pair.Key;
                dictext += ", ";
                dictext += pair.Value;
                dictext += "\r\n";

                //只显示前50条
                i++;
                if (i >= 50)
                    break;
            }

            MessageBox.Show(text: dictext,caption:"词库字典",buttons:MessageBoxButtons.OK);         

        }

        //判断输入的字符是否符合格式
        private bool CheckInputText(string inputStr)
        {
            //字符串不能为空
            if (string.IsNullOrEmpty(inputStr))
                return false;

            //判断字符串是否完全是中文
            //匹配中文字符的正则表达式
            //还要匹配换行符
            string patternCN = @"^[\u4e00-\u9fa5\r\n]+$";

            //匹配表达式，如果匹配成功，说明格式符合要求
            if (Regex.IsMatch(inputStr, patternCN))
            {
                return true;
            }
            else
                return false;
        }

        //将汉字转化为拼音
        private string ConvertCnToPinyin(string str)
        {
            string ConStr = "";
            //检测格式是否符合要求
            if (!CheckInputText(str))
            {
               return "要处理的内容必须是全中文";
            }

            //生成词典
            Entity.PinyinDictionary dict = new Entity.PinyinDictionary(@filename);

            //只取词典中的中文词汇，无需拼音
            List<string> wordList = dict.Dictionary.Keys.ToList<string>();

            //进行正向分词
            List<string> wordsLeft = Helper.Segmentation.SegMMLeftToRight(str, ref wordList);

            //判断分词是否正常返回
            if (wordsLeft == null)
            {               
                return "正向分词模块执行失败";
            }

            //转为拼音
            string pinyin = "";
            foreach (string word in wordsLeft)
            {
                //如果是单字，要检测字典中是否包含该单字
                if (word.Length == 1 && !dict.Dictionary.ContainsKey(word))
                {
                    //如果词典中不包含该中文单字，就要从微软的dll库读取拼音
                    pinyin = Helper.PinyinConvert.GetFirstPinYinCount(word.ToCharArray()[0]).ToLower();
                }
                else
                {
                    //一般情况不用检测，直接取词典中的拼音即可
                    pinyin = dict.Dictionary[word].ToLower();
                }

                //但如果不需要声调，还必须去掉声调
                if (!checkBoxWithTone.Checked)
                {
                    //这个正则表达式表示，去掉字符串中的数字
                    pinyin = Regex.Replace(pinyin, @"\d", "");
                }

                //将拼音显示出来
                ConStr += pinyin;

                //结尾加个空格
                ConStr += " ";
            }

            //去掉最后一个空格
            ConStr = ConStr.Trim();
            return ConStr;

        }

        //将汉字转化为拼音首字母
        public string ConvertCnToPinyinABC(string str)
        {
            string ConStr = "";
            //检测格式是否符合要求
            if (!CheckInputText(str))
            {
                return "要处理的内容必须是全中文";
            }

            //生成词典
            Entity.PinyinDictionary dict = new Entity.PinyinDictionary(@filename);

            //只取词典中的中文词汇，无需拼音
            List<string> wordList = dict.Dictionary.Keys.ToList<string>();

            //进行正向分词
            List<string> wordsLeft = Helper.Segmentation.SegMMLeftToRight(str, ref wordList);

            //判断分词是否正常返回
            if (wordsLeft == null)
            {
                return "正向分词模块执行失败";
            }

            //转为拼音
            string pinyin = "";
            foreach (string word in wordsLeft)
            {
                //如果是单字，要检测字典中是否包含该单字
                if (word.Length == 1 && !dict.Dictionary.ContainsKey(word))
                {
                    //如果词典中不包含该中文单字，就要从微软的dll库读取拼音
                  pinyin = Helper.PinyinConvert.GetFirstPinYinCount(word.ToCharArray()[0]).ToLower();
                }
                else
                {
                    //一般情况不用检测，直接取词典中的拼音即可
                    pinyin = dict.Dictionary[word].ToLower();
                }

                //但如果不需要声调，还必须去掉声调
                if (!checkBoxWithTone.Checked)
                {
                    //这个正则表达式表示，去掉字符串中的数字
                    pinyin = Regex.Replace(pinyin, @"\d", "");
                }

                //将拼音显示出来
                ConStr += pinyin;

                //结尾加个空格
                ConStr += " ";
            }

            //去掉最后一个空格
            ConStr = ConStr.Trim();     
      

            //取出首字母出来
            string ConStrABC="";
            string[] tempstr = ConStr.Split( );
            foreach (string tempstr1 in tempstr)
            {
                ConStrABC += tempstr1.Substring(0, 1);
            }

            return ConStrABC;
        }



        //测试读取词典
        private void buttonTestDict_Click(object sender, EventArgs e)
        {
            ReadDict();
        }

        //正向分词/逆向分词/双向分词  处理并导出到单元格的方法
        private void Segm(int seglb)
        {
            //测试分词
            //MessageBox.Show(SegWords(textBoxInput.Text));
            if (textBox1.Text == "")
            {
                MessageBox.Show("请先选择单元格区域");
                return;
            }

            string[,] brr = new string[arr1.GetUpperBound(0), arr1.GetUpperBound(1)];

            for (int i = 1; i <= arr1.GetUpperBound(0); i++)
            {
                for (int j = 1; j <= arr1.GetUpperBound(1); j++)
                    if (arr1[i, j] != null)
                    {
                        brr[i - 1, j - 1] = SegWords(arr1[i, j].ToString(), seglb);
                    }
            }
            
            this.Hide();
            try
            {
                Excel.Range rng = (Excel.Range)Globals.ThisAddIn.Application.InputBox(Prompt: "请选择存放区域，选择单个单元格即可", Type: 8);
                rng.Offset[0, 0].Resize[brr.GetUpperBound(0) + 1, brr.GetUpperBound(1)+1].Value = brr;
                this.Show();
            }
            catch
            {
                this.Show();
                return;
            }


        }



        //正向分词/逆向分词/双向分词  统计 处理并导出到单元格的方法
        private void SegWordsStatis(int seglb)
        {
            //测试分词
            //MessageBox.Show(SegWords(textBoxInput.Text));
            if (textBox1.Text == "")
            {
                MessageBox.Show("请先选择单元格区域");
                return;
            }

			string SegWordsStatis ="";         

            for (int i = 1; i <= arr1.GetUpperBound(0); i++)
            {
                for (int j = 1; j <= arr1.GetUpperBound(1); j++)
                    if (arr1[i, j] != null)
                    {
                        SegWordsStatis += SegWords(arr1[i, j].ToString(), seglb).Substring(5);//去掉分词那几个字；
                    }
            }

            //去掉最后一个逗号
            SegWordsStatis = SegWordsStatis.Substring(0, SegWordsStatis.Length - 1);
			
			 //字典，来统计
            Dictionary<string, int> dic = new Dictionary<string, int>();
            

            string[] tempstr = SegWordsStatis.Split(',');
            foreach (string tempstr1 in tempstr)
            {
               	 if(dic.ContainsKey(tempstr1))
				 {
					dic[tempstr1] += 1;
				 }
				 else
				 {
					dic[tempstr1] = 1;
				 }		
            }

            //把字典内容循环写出来
            string[,] crr = new string[dic.Count+1,2];
            int tempi = 0;
            crr[0, 0] = "词频";
            crr[0, 1] = "频数";
   
            foreach (KeyValuePair<string, int> kvp in dic)
            {
                tempi = tempi + 1;
                crr[tempi, 0] = kvp.Key;
                crr[tempi, 1] = kvp.Value.ToString();                
            }
                       		
            //导出到单元格
            this.Hide();
            try
            {
                Excel.Range rng = (Excel.Range)Globals.ThisAddIn.Application.InputBox(Prompt: "请选择存放区域，选择单个单元格即可", Type: 8);
                rng.Offset[0, 0].Resize[crr.GetUpperBound(0) + 1, crr.GetUpperBound(1)+1].Value = crr;
                this.Show();
            }
            catch
            {
                this.Show();
                return;
            }


        }



        //转换中文为拼音
        private void buttonCnToPinyin_Click(object sender, EventArgs e)
        {
            //转为拼音
          // MessageBox.Show(ConvertCnToPinyin("俊哥哥知道重庆很重要"));
            if (textBox1.Text == "")
            {
                MessageBox.Show("请先选择单元格区域");
                return;
            }

            string[,] brr = new string[arr1.GetUpperBound(0), arr1.GetUpperBound(1)];

            for (int i = 1; i <= arr1.GetUpperBound(0); i++)
            {
                for (int j = 1; j <= arr1.GetUpperBound(1); j++)
                    if (arr1[i, j] != null)
                    {
                        brr[i-1, j-1] = ConvertCnToPinyin(arr1[i, j].ToString());
                    }
            }          

            this.Hide();
            try
            {
                Excel.Range rng = (Excel.Range)Globals.ThisAddIn.Application.InputBox(Prompt: "请选择存放区域，选择单个单元格即可", Type: 8);
                rng.Offset[0, 0].Resize[brr.GetUpperBound(0) + 1, brr.GetUpperBound(1) + 1].Value =brr;
                this.Show();
            }
            catch
            {
                this.Show();
                return;
            }

        }

   
        private string  CleanNoCh(string str)
        {
            return Regex.Replace(str, @"[^\u4e00-\u9fa5]*", "");        
        }



        //选择用户自己的词典
        private void buttonSelectDict_Click(object sender, EventArgs e)
        {
            //将用户选择的文件名和路径，指定为用户词典文件
            if (openFileDialogDict.ShowDialog() == DialogResult.OK)
            {
                filename = openFileDialogDict.FileName;
                MessageBox.Show("词典更改为：" + filename);
            }
        }



        private void button1_Click(object sender, EventArgs e)
        {        
            
            
          //  this.Visible = false;
            this.Hide();           
            try
            {
                rng1 = (Excel.Range)Globals.ThisAddIn.Application.InputBox(Prompt: "请选择数据区域", Type: 8);
            }
            catch
            {
                return;
            }


            try
            {
                rng1 = Globals.ThisAddIn.Application.Intersect(sht.UsedRange, rng1);
                if (rng1.Address == "$A$1" || rng1 == null)
                {
                    MessageBox.Show(text: "请不要选择空白区域", buttons: MessageBoxButtons.OK, caption: "");
                    return;
                }

                if (rng1.Count == 1)
                {
                    MessageBox.Show(text: "请不要选择单个单元格", buttons: MessageBoxButtons.OK, caption: "");
                    return;
                }
            }
            catch
            {
                MessageBox.Show(text: "请不要选择空白区域", buttons: MessageBoxButtons.OK, caption: "",icon:MessageBoxIcon.Warning);
                return;
            }
            arr1 = rng1.Value;
            textBox1.Text = rng1.Address;
            this.Show();

           // MessageBox.Show(arr1.GetUpperBound(0).ToString());
            //MessageBox.Show(arr1.GetUpperBound(1).ToString());
        }

        private void buttonSegm_Click(object sender, EventArgs e)
        {
            //正向分词
            Segm(1);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //逆向分词
            Segm(2);
        }

        private void button3_Click(object sender, EventArgs e)
        {
           //双向分词
            Segm(3);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //转为拼音首字母
            // MessageBox.Show(ConvertCnToPinyin("俊哥哥知道重庆很重要"));
            if (textBox1.Text == "")
            {
                MessageBox.Show("请先选择单元格区域");
                return;
            }

            string[,] brr = new string[arr1.GetUpperBound(0), arr1.GetUpperBound(1)];

            for (int i = 1; i <= arr1.GetUpperBound(0); i++)
            {
                for (int j = 1; j <= arr1.GetUpperBound(1); j++)
                    if (arr1[i, j] != null)
                    {
                        brr[i - 1, j - 1] = ConvertCnToPinyinABC(arr1[i, j].ToString());
                    }
            }

            //this.Visible = false;
            this.Hide();
            try
            {
                Excel.Range rng = (Excel.Range)Globals.ThisAddIn.Application.InputBox(Prompt: "请选择存放区域，选择单个单元格即可", Type: 8);
                rng.Offset[0, 0].Resize[brr.GetUpperBound(0) + 1, brr.GetUpperBound(1) + 1].Value = brr;
                this.Show();
            }
            catch
            {
                this.Show();
                return;
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            SegWordsStatis(1);
        }

    
     
      
             
    }
}
