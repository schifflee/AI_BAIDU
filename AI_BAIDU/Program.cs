using Baidu.Aip.Nlp;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AI_BAIDU
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] paths = Directory.GetFiles(@"QuestionLibraries\", "*.xls");
            /* 短文本相似度,最大512字节（256字符）,
             CNN（卷积神经网络）模型
             模型语义泛化能力介于 BOW / RNN 之间，对序列输入敏感，相较于 GRNN 模型的一个显著优点是计算效率会更高些。
             返回示例
                 {
                     "log_id": 12345,
                     "texts":{
                         "text_1":"浙富股份",
                         "text_2":"万事通自考网"
                     },
                     "score":0.3300237655639648 //相似度结果
                 },
              var result = client.Simnet(str1, str2);
             短文本相似度,如果有可选参数,带参数调用短文本相似度
              var options = new Dictionary<string, object> { { "model", "GRNN" } };
              带参数调用短文本相似度
             result = client.Simnet(text1, text2, options);*/
            int count, allcount = 0;
            foreach (string path in paths)
            {
                util.Simnet(path, @"Simnet\", @"Simnet\SimnetError.txt", out count);
                allcount += count;
            }
            Console.WriteLine("共完成短文本相似度计算{0}条。", allcount);

            /*** 2020年4月23日桐城兴尔旺
             * 下面部分是短文本纠错的内容，
            //用于保存短文本纠错结果
            string resultpath = ConfigurationManager.AppSettings["ResultPath"];
            int count, allcount = 0;
            foreach (string path in paths)
            {
                //进行文本纠错
                util.Ecnet(path, resultpath, out count);
                allcount += count;
                Console.WriteLine("本次完成纠错{0}条，共完成纠错{1}条。", count, allcount);
            }
            Console.WriteLine("总计完成纠错{0}条", allcount);
            ***/

            Console.ReadLine();
        }


    }
}
