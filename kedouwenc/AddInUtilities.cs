using System.Data;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


//Calling Code in an Application-Level Add-in from VBA 
//从 VBA 调用 VSTO 外接程序中的代码
//https://msdn.microsoft.com/zh-cn/library/bb608614.aspx
//VBAd代码
//Sub CallVSTOMethod()
//     Dim addIn As COMAddIn
//     Dim automationObject As Object
//     Set addIn = Application.COMAddIns("kedouwenc")
//     Set automationObject = addIn.Object
//     automationObject.ImportData
//End Sub

namespace kedouwenc
{
        [ComVisible(true)]
        public interface IAddInUtilities
        {
            void ImportData();
        }

        [ComVisible(true)]
        [ClassInterface(ClassInterfaceType.None)]
        public class AddInUtilities : IAddInUtilities
        {
            // This method tries to write a string to cell A1 in the active worksheet.
            public void ImportData()
            {
                Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

                if (activeWorksheet != null)
                {
                    Excel.Range range1 = activeWorksheet.get_Range("A1", System.Type.Missing);
                    range1.Value2 = "This is my data";
                }
            }
        }
    
}
