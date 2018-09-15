using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Baidu.Aip.Ocr;

namespace kedouwenc
{
    class BaiduAI_InterActive
    {


        
        public Baidu.Aip.Ocr.Ocr baidu_ai_InterActive() {

            // 设置APPID/AK/SK
            string APP_ID = "11430855";
            string API_KEY = "YD56VqTWdFzN0ZdhWa2nngOg";
            string SECRET_KEY = "ZphWGfObsRjjv1MrEyurY8h8WCf8geSy";

            var client = new Baidu.Aip.Ocr.Ocr(API_KEY, SECRET_KEY);
            client.Timeout = 60000;  // 修改超时时间

            return client;
        }


    }
}
