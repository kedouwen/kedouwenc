﻿using System;
//using System.Collections.Generic;
//using System.Linq;
using System.Text;
using Microsoft.International.Converters.PinYinConverter;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;

namespace kedouwenc.Helper
{
    static class PinyinConvert
    {
        /// <summary>
        /// 返回单个简体中文字的拼音列表
        /// </summary>
        /// <param name="inputChar">简体中文单字</param>      
        public static ReadOnlyCollection<string> GetPinYinWithTone(Char inputChar)
        {
            string patternCN = @"^[\u4e00-\u9fa5\r\n]+$";
            if (! Regex.IsMatch(inputChar.ToString(), patternCN))
            {
                return new ReadOnlyCollection<string>(inputChar.ToString().Split(' '));
            }
            else {
                ChineseChar chineseChar = new ChineseChar(inputChar);
                return chineseChar.Pinyins;
            }

           
        }

        /// <summary>
        /// 返回单个简体中文字的拼音个数
        /// </summary>
        /// <param name="inputChar">简体中文单字</param>      
        public static short GetPinYinCount(Char inputChar)
        {
            ChineseChar chineseChar = new ChineseChar(inputChar);
            return chineseChar.PinyinCount;
        }

        /// <summary>
        /// 返回单个简体中文字拼音列表中的第一个拼音
        /// </summary>
        /// <param name="inputChar">简体中文单字</param>      
        public static string GetFirstPinYinCount(Char inputChar)
        {
            //得到第一个拼音
            return GetPinYinWithTone(inputChar)[0];
        }

      

    }
}
