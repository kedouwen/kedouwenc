using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace kedouwenc
{
    public static class RangeSelectHelper
    {

        public static void blankrange()
        {
            MessageBox.Show("请不要选择空白区域", "友情提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void textrange()
        {
            MessageBox.Show(text: "选择的单元格没有文本区域。", caption: "提示", buttons: MessageBoxButtons.OK);
        }

        public static void valuerange()
        {
            MessageBox.Show("选择的单元格没有数值区域", "友情提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void formularange()
        {
            MessageBox.Show(text: "选择的单元格没有公式。", caption: "提示", buttons: MessageBoxButtons.OK);
        }

        public static void jsontorange(Newtonsoft.Json.Linq.JObject result)
        {
            Excel.Worksheet sht = Globals.ThisAddIn.Application.ActiveSheet;
            try
            {
                Excel.Range rng = (Excel.Range)Globals.ThisAddIn.Application.InputBox(Prompt: "请选择存放区域，选择单个单元格即可", Type: 8);
                

                string[] resultwords = new string[(int)result["words_result_num"]];
                int i = 0;
                foreach (var word in result["words_result"])
                {                    
                    resultwords[i] = (string)word["words"];
                    i++;
                }

                rng.Offset[0, 0].Resize[resultwords.Length, 1].Value = Globals.ThisAddIn.Application.WorksheetFunction.Transpose(resultwords);

            }
            catch
            {
                MessageBox.Show("选择单元格！");
            }
            



        }





    }

       
}
