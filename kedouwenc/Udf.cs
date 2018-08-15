//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using Office = Microsoft.Office.Core;
//using Microsoft.Office.Interop.Excel;
//using Excel = Microsoft.Office.Interop.Excel;

//namespace kedouwenc
//{
//    class Udf
//    {
//        public bool IsNumber(object expression)
//        {
//            bool IsNumber;
//            double retnum;
//            IsNumber = Double.TryParse(Convert.ToString(expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retnum);
//            return IsNumber;
//        }








       
//    }
//}
using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;


namespace AutomationAddin
{


    // Replace the Guid below with your own guid that

    // you generate using Create GUID from the Tools menu

    [Guid("5268ABE2-9B09-439d-BE97-2EA60E103EF6")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    public class MyFunctions
    {
        public MyFunctions()
        {

        }

        public string GetStars(double Number)
        {
            string s = "";
            for (double i = 0; i < Number; i++)
            {
                s = s + "*";
            }
            return s;
        }

        public double MultiplyNTimes(double number1, double number2, double timesToMultiply)
        {
            double result = number1;
            for (double i = 0; i < timesToMultiply; i++)
            {
                result = result * number2;
            }


            return result;
        }


        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type type)
        {


            Registry.ClassesRoot.CreateSubKey(

              GetSubKeyName(type, "Programmable"));

            RegistryKey key = Registry.ClassesRoot.OpenSubKey(

              GetSubKeyName(type, "InprocServer32"), true);

            key.SetValue("",

              System.Environment.SystemDirectory + @"\mscoree.dll",

              RegistryValueKind.String);
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type type)
        {


            Registry.ClassesRoot.DeleteSubKey(

              GetSubKeyName(type, "Programmable"), false);
        }

        private static string GetSubKeyName(Type type,

          string subKeyName)
        {

            System.Text.StringBuilder s =

              new System.Text.StringBuilder();

            s.Append(@"CLSID\{");

            s.Append(type.GUID.ToString().ToUpper());

            s.Append(@"}\");

            s.Append(subKeyName);

            return s.ToString();

        }
    }
}


