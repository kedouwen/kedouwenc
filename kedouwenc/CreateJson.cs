using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System.IO;

namespace kedouwenc
{
    public partial class CreateJson : UserControl
    {
        public CreateJson()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Range rng;
            object[,] arr;
            string temptxt = "";
            string tempjson = "";

            rng = Globals.ThisAddIn.Application.Selection;
            arr = rng.Value;           

            for (int i = 2; i <= arr.GetUpperBound(0); i++)
            {
                temptxt = "";
                for (int j = 1; j <= arr.GetUpperBound(1); j++)
                {
                    temptxt = temptxt + arr[1, j].ToString().ToLower() + ": '" + arr[i, j] + "',";
                }
                temptxt =  temptxt.Substring(0, temptxt.Length - 2);
                tempjson = tempjson + " {        " + temptxt + "'},";
            }


            tempjson = "[" + tempjson.Substring(0, tempjson.Length - 1) + "]";
            //MessageBox.Show(tempjson);

            textBox1.Text = ConvertJsonString(tempjson) + "\r\n";

            if (textBox1.Text != "")
                Clipboard.SetDataObject(textBox1.Text);



        }

        private string ConvertJsonString(string str)
        {
            //格式化json字符串
            JsonSerializer serializer = new JsonSerializer();
            TextReader tr = new StringReader(str);
            JsonTextReader jtr = new JsonTextReader(tr);
            object obj = serializer.Deserialize(jtr);
            if (obj != null)
            {
                StringWriter textWriter = new StringWriter();
                JsonTextWriter jsonWriter = new JsonTextWriter(textWriter)
                {
                    Formatting = Formatting.Indented,
                    Indentation = 4,
                    IndentChar = ' '
                };
                serializer.Serialize(jsonWriter, obj);
                return textWriter.ToString();
            }
            else
            {
                return str;
            }
        }



    }
}
