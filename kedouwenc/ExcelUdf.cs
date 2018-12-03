using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace kedouwenc
{


    // Replace the Guid below with your own guid that

    // you generate using Create GUID from the Tools menu

    [Guid("311B7A43-A33C-463D-A332-17F60DD1B950")]   
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    public class MyFunctions
    {
        public MyFunctions()
        {

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


        public string GetStars(double Number)
        {
            string s = "";
            for (double i = 0; i < Number; i++)
            {
                s = s + "*";
            }
            return s;
        }


        public double AddNumbers(double Number1, [Optional] object Number2, [Optional] object Number3)
        {
            double result = 0;
            result += Convert.ToDouble(Number1);


            if (!(Number2 is System.Reflection.Missing))
            {
                Excel.Range r2 = Number2 as Excel.Range;
                double d2 = Convert.ToDouble(r2.Value2);
                result += d2;
            }


            if (!(Number3 is System.Reflection.Missing))
            {
                Excel.Range r3 = Number3 as Excel.Range;
                double d3 = Convert.ToDouble(r3.Value2);
                result += d3;
            }


            return result;
        }


        public double CalculateArea(object Range)
        {
            Excel.Range r = Range as Excel.Range;
            return Convert.ToDouble(r.Width) * Convert.ToDouble(r.Height);
        }


        public double NumberOfCells(object Range)
        {
            Excel.Range r = Range as Excel.Range;
            return r.Cells.Count;
        }


        public string ToUpperCase(string input)
        {
            return input.ToUpper();
        }




    }
}