using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Baidu.Aip.Ocr;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

//Install-Package Baidu.AI



namespace kedouwenc
{
    class BaiduAI
    {

        public void GeneralBasicDemo()
        {
            BaiduAI_InterActive ai_InterActive = new BaiduAI_InterActive();
            var client=ai_InterActive.baidu_ai_InterActive();

            //var image = File.ReadAllBytes(@"D:\666.jpg");

            byte[] image=null;
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "*jpg|*.JPG|*.GIF|*.GIF|*.BMP|*.BMP";
            DialogResult fdresult = fileDialog.ShowDialog();
            if (fdresult == DialogResult.OK)
            {
                image = File.ReadAllBytes(fileDialog.FileName);
            }
            
            Excel.Worksheet actsheet = Globals.ThisAddIn.Application.ActiveSheet;            
            actsheet.Shapes.AddPicture(fileDialog.FileName,Microsoft.Office.Core.MsoTriState.msoFalse,
                Microsoft.Office.Core.MsoTriState.msoCTrue,100,100,100,100);


            // 调用通用文字识别, 图片参数为本地图片，可能会抛出网络等异常，请使用try/catch捕获
            var result = client.GeneralBasic(image);
            Console.WriteLine(result);
            // 如果有可选参数
            var options = new Dictionary<string, object>{
            {"language_type", "CHN_ENG"},
            {"detect_direction", "true"},
            {"detect_language", "true"},
            {"probability", "false"} //行置信度信息；如果输入参数 probability = true 则输出
        };
            // 带参数调用通用文字识别, 图片参数为本地图片
            result = client.GeneralBasic(image, options);
            Console.WriteLine(result);
            RangeSelectHelper.jsontorange(result);



        }


        public void AccurateBasicDemo()
        {
            BaiduAI_InterActive ai_InterActive = new BaiduAI_InterActive();
            var client = ai_InterActive.baidu_ai_InterActive();                    

            byte[] image = null;
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "*jpg|*.JPG|*.GIF|*.GIF|*.BMP|*.BMP";
            DialogResult fdresult = fileDialog.ShowDialog();
            if (fdresult == DialogResult.OK)
            {
                image = File.ReadAllBytes(fileDialog.FileName);
            }

            Excel.Worksheet actsheet = Globals.ThisAddIn.Application.ActiveSheet;
            actsheet.Shapes.AddPicture(fileDialog.FileName, Microsoft.Office.Core.MsoTriState.msoFalse,
            Microsoft.Office.Core.MsoTriState.msoCTrue, 100, 100, 100, 100);
            // 如果有可选参数
            var options = new Dictionary<string, object>{
                 {"detect_direction", "false"},
                    {"probability", "false"} };
            // 带参数调用通用文字识别（高精度版）
            var result = client.AccurateBasic(image, options);
            Console.WriteLine(result);
            RangeSelectHelper.jsontorange(result);
        }



        public void IdcardDemo()
        {

            BaiduAI_InterActive ai_InterActive = new BaiduAI_InterActive();
            var client = ai_InterActive.baidu_ai_InterActive();

            var image = File.ReadAllBytes(@"D:\666.jpg");
            var idCardSide = "front";

            // 调用身份证识别，可能会抛出网络等异常，请使用try/catch捕获
            var result = client.Idcard(image, idCardSide);
            Console.WriteLine(result);
            // 如果有可选参数
            var options = new Dictionary<string, object>{
                {"detect_direction", "true"},
                {"detect_risk", "true"}
            };
            // 带参数调用身份证识别
            result = client.Idcard(image, idCardSide, options);
            Console.WriteLine(result);
        }


        public void TableRecognitionRequestDemo()
        {
            BaiduAI_InterActive ai_InterActive = new BaiduAI_InterActive();
            var client = ai_InterActive.baidu_ai_InterActive();
            var image = File.ReadAllBytes(@"D:\666.jpg");
            // 调用表格文字识别，可能会抛出网络等异常，请使用try/catch捕获
            var result = client.TableRecognitionRequest(image);
            Console.WriteLine(result);
        }


        public void TableRecognitionGetResultDemo()
        {
            BaiduAI_InterActive ai_InterActive = new BaiduAI_InterActive();
            var client = ai_InterActive.baidu_ai_InterActive();
            //var image = File.ReadAllBytes(@"D:\666.jpg");
           // var result_get = client.TableRecognitionRequest(image);
            //MessageBox.Show(result_get["result"][0]["request_id"].ToString());
           // var requestId = result_get["result"][0]["request_id"].ToString();

            var requestId = "11430855_592019";

            // 调用表格识别结果，可能会抛出网络等异常，请使用try/catch捕获
            var result = client.TableRecognitionGetResult(requestId);
            Console.WriteLine(result);
            System.Diagnostics.Debug.WriteLine(result);


            MessageBox.Show(result.ToString());
            // 如果有可选参数
            var options = new Dictionary<string, object>{{"result_type", "excel"}    };
            // 带参数调用表格识别结果
            result = client.TableRecognitionGetResult(requestId, options);
            Console.WriteLine(result);
        }




    }
}
