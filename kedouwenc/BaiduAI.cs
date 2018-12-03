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




        public byte[] selectimage() {
            byte[] image = null;
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "*jpg|*.JPG|*.GIF|*.GIF|*.BMP|*.BMP";
            DialogResult fdresult = fileDialog.ShowDialog();
            if (fdresult == DialogResult.OK)
            {
                image = File.ReadAllBytes(fileDialog.FileName);
            }
            return image;

        }

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



        public string TableRecognitionRequestDemo()
        {
            BaiduAI_InterActive ai_InterActive = new BaiduAI_InterActive();
            var client = ai_InterActive.baidu_ai_InterActive();
            //var image = File.ReadAllBytes(@"D:\666.jpg");

            byte[] image = null;
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "*jpg|*.JPG|*.GIF|*.GIF|*.BMP|*.BMP";
            DialogResult fdresult = fileDialog.ShowDialog();
            if (fdresult == DialogResult.OK)
            {
                image = File.ReadAllBytes(fileDialog.FileName);
            }

            //Excel.Worksheet actsheet = Globals.ThisAddIn.Application.ActiveSheet;
            //actsheet.Shapes.AddPicture(fileDialog.FileName, Microsoft.Office.Core.MsoTriState.msoFalse,
            //Microsoft.Office.Core.MsoTriState.msoCTrue, 100, 100, 100, 100);


            // 调用表格文字识别，可能会抛出网络等异常，请使用try/catch捕获
            var result = client.TableRecognitionRequest(image);
            //Console.WriteLine(result);
            return result["result"][0]["request_id"].ToString();
        }


        public void TableRecognitionGetResultDemo()
        {
            BaiduAI_InterActive ai_InterActive = new BaiduAI_InterActive();
            var client = ai_InterActive.baidu_ai_InterActive();

            var requestId = TableRecognitionRequestDemo();
            //var requestId = "11430855_597217";
            

            //byte[] image = null;
            //OpenFileDialog fileDialog = new OpenFileDialog();
            //fileDialog.Filter = "*jpg|*.JPG|*.GIF|*.GIF|*.BMP|*.BMP";
            //DialogResult fdresult = fileDialog.ShowDialog();
            //if (fdresult == DialogResult.OK)
            //{
            //    image = File.ReadAllBytes(fileDialog.FileName);
            //}
            //var result_GET = client.TableRecognitionRequest(image);

            //var requestId = result_GET["result"][0]["request_id"].ToString();

            Newtonsoft.Json.Linq.JObject result;
            
            int condition = 0;
            /* do 循环执行 */
            do
            {
                // 调用表格识别结果，可能会抛出网络等异常，请使用try/catch捕获
                result = client.TableRecognitionGetResult(requestId);
                //System.Diagnostics.Debug.WriteLine(result["result"]["ret_code"].ToString());
                try
                {
                    condition = (int)result["result"]["ret_code"];
                }
                catch
                {
                    condition = 99999;
                }
                System.Diagnostics.Debug.WriteLine(condition);


            } while (condition != 3);
                      
            //Console.WriteLine(result);
            System.Diagnostics.Debug.WriteLine(result);           
           
            // 如果有可选参数
            //var options = new Dictionary<string, object>{{"result_type", "excel"}    };
            // 带参数调用表格识别结果
            //result = client.TableRecognitionGetResult(requestId, options);

            MessageBox.Show(result.ToString());

            Excel.Range rng = (Excel.Range)Globals.ThisAddIn.Application.InputBox(Prompt: "请选择存放区域，选择单个单元格即可", Type: 8);

            string result_file_url = (string)result["result"]["result_data"];

            rng.Offset[0, 0].Resize[1, 1].Value = result_file_url;






        }




    }
}
