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
    }

       
}
