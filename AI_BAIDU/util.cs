using Baidu.Aip.Nlp;
using Newtonsoft.Json.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace AI_BAIDU
{
    class util
    {
        /// <summary>
        /// 短文本相似度计算，主要程序
        /// </summary>
        /// <param name="path">excel试题位置</param>
        /// <param name="ResultPath">存放excel结果的文件夹</param>
        /// <param name="count">共进行了多少次比较</param>
        public static void Simnet(string ExcelPath, string ResultPath, string ErrorPath, out int count)
        {
            count = 0;
            Nlp client = util.CreateClient();

            FileStream ResultFileStream = new FileStream(ResultPath + Path.GetFileName(ExcelPath), FileMode.OpenOrCreate, FileAccess.ReadWrite);
            //每一科放在一个工作薄中
            HSSFWorkbook workbook = new HSSFWorkbook();
            StreamWriter ErrorSW = new StreamWriter(ErrorPath, true);
            DataTable dt = util.ExcelToDataTable(ExcelPath, true);
            List<Question> alllist = util.DatatableConvertToQuestion(dt);
            Dictionary<string, List<Question>> DICQuestionSEPChapter = new Dictionary<string, List<Question>>();
            var options = new Dictionary<string, object> { { "model", "GRNN" } };
            ///将各数据按章节分配到QuestionSEPChapter
            foreach (Question q in alllist)
            {
                if (DICQuestionSEPChapter.ContainsKey(q.Chapter.Trim()))
                {
                    DICQuestionSEPChapter[q.Chapter.Trim()].Add(q);
                }
                else
                {
                    List<Question> list = new List<Question>();
                    list.Add(q);
                    DICQuestionSEPChapter.Add(q.Chapter.Trim(), list);
                }
            }
            ///下面部分开始按章节计算相似度
            foreach (string key in DICQuestionSEPChapter.Keys)
            {
                int chapterCount = 1;
                List<SimnetResult> ListsimnetResult = new List<SimnetResult>();
                //每一章放在一个工作表中
                ISheet sheet = workbook.CreateSheet(key);
                for (int i = 0; i < DICQuestionSEPChapter[key].Count - 1; i++)
                {
                    Question questionA = DICQuestionSEPChapter[key][i];
                    string middlewareA = new Regex("[_]{5,}").Replace(questionA.Title, util.GetAnswerStr(questionA, questionA.Answer));
                    String textA = middlewareA.Length > 255 ? middlewareA.Substring(0, 255) : middlewareA;
                    for (int j = i + 1; j < DICQuestionSEPChapter[key].Count; j++)
                    {
                        Question questionB = DICQuestionSEPChapter[key][j];
                        string middlewareB = new Regex("[_]{5,}").Replace(questionB.Title, util.GetAnswerStr(questionB, questionA.Answer));
                        string textB = middlewareB.Length > 255 ? middlewareB.Substring(0, 255) : middlewareB;
                        try
                        {
                            JObject result = client.Simnet(textA, textB, options);
                            JToken error_code;
                            //出现错误
                            if (result.TryGetValue("error_code", out error_code))
                            {
                                ErrorSW.WriteLine("subject:{0}||questionA:{1}||questionB{2}||error_code:{3};", questionA.Subject, questionA.SNID, questionB.SNID, error_code.ToString());
                                Console.WriteLine("error_code:" + error_code);
                            }
                            else
                            {
                                SimnetResult simnetResult = new SimnetResult();
                                simnetResult.Id = chapterCount;
                                simnetResult.Subject = questionA.Subject;
                                simnetResult.SNIDA = questionA.SNID;
                                simnetResult.SNIDB = questionB.SNID;
                                simnetResult.TitleA = questionA.Title;
                                simnetResult.TitleB = questionB.Title;
                                simnetResult.Score = Convert.ToDecimal(result["score"]);
                                simnetResult.IssubA = middlewareA.Length > 255 ? true : false;
                                simnetResult.IssubB = middlewareB.Length > 255 ? true : false;
                                ListsimnetResult.Add(simnetResult);
                                Console.WriteLine("key:{0}, chaptercount:{1}, score: {2}, issubA:{3}, issubB:{4}", key, chapterCount, simnetResult.Score, simnetResult.IssubA, simnetResult.IssubB);
                                chapterCount++;
                            }
                            Thread.Sleep(500);
                        }
                        catch (Exception e)
                        {
                            ErrorSW.WriteLine(e.Message);
                            Thread.Sleep(500);
                            continue;
                        }
                    }
                }
                InsertIntoExcelsheet(sheet, ListsimnetResult);
                workbook.Write(ResultFileStream);
                count += chapterCount;
            }
        }
        /// <summary>
        /// 将计算结果写入到EXCEL工作表中
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="listsimnetResult"></param>
        private static ISheet InsertIntoExcelsheet(ISheet sheet, List<SimnetResult> listsimnetResult)
        {
            //利用反射生成标题行
            IRow headRow = sheet.CreateRow(0);
            PropertyInfo[] propertyInfos = typeof(SimnetResult).GetProperties();
            for (int i = 0; i < propertyInfos.Length; i++)
            {
                headRow.CreateCell(i).SetCellValue(propertyInfos[i].Name);
            }
            for (int j = 0; j < listsimnetResult.Count; j++)
            {
                IRow row = sheet.CreateRow(j + 1);
                for (int k = 0; k < headRow.Cells.Count; k++)
                {
                    ICell cell = row.CreateCell(k);
                    cell.SetCellValue(typeof(SimnetResult).GetProperty(propertyInfos[k].Name).GetValue(listsimnetResult[j]).ToString());
                }
            }
            return sheet;
        }


        /// <summary>
        /// 文本纠错，主要程序
        /// </summary>
        /// <param name="path">excel试题库</param>
        /// <param name="resultpath">错误内容存放的位置</param>
        public static void Ecnet(string path, string resultpath, out int count)
        {
            count = 0;
            DataTable dt = util.ExcelToDataTable(path, true);
            List<Question> list = util.DatatableConvertToQuestion(dt);
            Nlp client = util.CreateClient();
            StringBuilder sb = null;
            StreamWriter errorprint = new StreamWriter("error.txt", true);
            StreamWriter ResultPrint = new StreamWriter(resultpath, true);
            foreach (Question question in list)
            {
                sb = new StringBuilder(question.Title.Replace("_______", util.GetAnswerStr(question, question.Answer)));
                sb.Append(question.Choosea);
                sb.Append(question.Chooseb);
                sb.Append(question.Choosec);
                sb.Append(question.Choosed);
                sb.Append(question.Explain);
                int i = 0;
                while (sb.Length - 255 * i > 0)
                {
                    try
                    {
                        JObject result = client.Ecnet(sb.Length - 255 * i > 255 ? sb.ToString().Substring(i * 255, 255) : sb.ToString().Substring(i * 255, sb.Length - 255 * i));
                        JToken error_code;
                        //如果发生错误进行记录
                        if (result.TryGetValue("error_code", out error_code))
                        {
                            string ErrorPrint = question.AllID + "||" + error_code.ToString() + "||" + result["error_msg"];
                            errorprint.WriteLine(ErrorPrint);
                            errorprint.Flush();
                            Thread.Sleep(500);
                            break;
                        }
                        else //如果获得正确结果的处理
                        {
                            decimal score = Convert.ToDecimal(result["item"]["score"]);
                            string CorrectPrint = "Count:" + count + "||path:" + path + "||ALLID:" + question.AllID + "||SNID:" + question.SNID + "||log_id:" + result["log_id"];
                            Console.WriteLine(CorrectPrint);
                            //如果有需要纠错的内容
                            if (score != 0)
                            {
                                ResultPrint.WriteLine(CorrectPrint + result.ToString());
                                ResultPrint.Flush();
                                Console.WriteLine(result);
                            }
                        }
                        i++;
                        Thread.Sleep(500);
                    }
                    catch (Exception e)
                    {
                        errorprint.WriteLine(e.Message); errorprint.Flush();
                    }
                }
                count++;
            }
            errorprint.Close();
            ResultPrint.Close();
        }

        /// <summary>
        /// 将试题的ABCD的答案转化为对应选项的内容
        /// </summary>
        /// <param name="q"></param>
        /// <param name="answer"></param>
        /// <returns></returns>
        public static string GetAnswerStr(Question q, string answer)
        {
            switch (answer.Trim().ToLower())
            {
                case "a": return q.Choosea;
                case "b": return q.Chooseb;
                case "c": return q.Choosec;
                case "d": return q.Choosed;
                default: return null;
            }
        }

        /// <summary>
        /// 获得配置文件生成Client对象
        /// </summary>
        /// <returns></returns>
        public static Nlp CreateClient()
        {
            // 设置APPID/AK/SK
            string APP_ID = ConfigurationManager.AppSettings["APP_ID"];
            string API_KEY = ConfigurationManager.AppSettings["API_KEY"];
            string SECRET_KEY = ConfigurationManager.AppSettings["SECRET_KEY"];
            Baidu.Aip.Nlp.Nlp client = new Baidu.Aip.Nlp.Nlp(API_KEY, SECRET_KEY);
            client.Timeout = 60000;
            return client;
        }

        /// <summary>
        /// 将excel导入到datatable
        /// </summary>
        /// <param name="filePath">excel路径</param>
        /// <param name="isColumnName">第一行是否是列名</param>
        /// <returns>返回datatable</returns>
        public static DataTable ExcelToDataTable(string filePath, bool isColumnName)
        {
            DataTable dataTable = null;
            FileStream fs = null;
            DataColumn column = null;
            DataRow dataRow = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            ICell cell = null;
            int startRow = 0;
            try
            {
                using (fs = File.OpenRead(filePath))
                {
                    // 2007版本
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    // 2003版本
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {
                        sheet = workbook.GetSheetAt(0);//读取第一个sheet，当然也可以循环读取每个sheet
                        dataTable = new DataTable();
                        if (sheet != null)
                        {
                            int rowCount = sheet.LastRowNum;//总行数
                            if (rowCount > 0)
                            {
                                IRow firstRow = sheet.GetRow(0);//第一行
                                int cellCount = firstRow.LastCellNum;//列数

                                //构建datatable的列
                                if (isColumnName)
                                {
                                    startRow = 1;//如果第一行是列名，则从第二行开始读取
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                    {
                                        cell = firstRow.GetCell(i);
                                        if (cell != null)
                                        {
                                            if (cell.StringCellValue != null)
                                            {
                                                column = new DataColumn(cell.StringCellValue);
                                                dataTable.Columns.Add(column);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                    {
                                        column = new DataColumn("column" + (i + 1));
                                        dataTable.Columns.Add(column);
                                    }
                                }

                                //填充行
                                for (int i = startRow; i <= rowCount; ++i)
                                {
                                    row = sheet.GetRow(i);
                                    if (row == null) continue;

                                    dataRow = dataTable.NewRow();
                                    for (int j = row.FirstCellNum; j < cellCount; ++j)
                                    {
                                        cell = row.GetCell(j);
                                        if (cell == null)
                                        {
                                            dataRow[j] = "";
                                        }
                                        else
                                        {
                                            //CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)
                                            switch (cell.CellType)
                                            {
                                                case CellType.Blank:
                                                    dataRow[j] = "";
                                                    break;
                                                case CellType.Numeric:
                                                    short format = cell.CellStyle.DataFormat;
                                                    //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理
                                                    if (format == 14 || format == 31 || format == 57 || format == 58)
                                                        dataRow[j] = cell.DateCellValue;
                                                    else
                                                        dataRow[j] = cell.NumericCellValue;
                                                    break;
                                                case CellType.String:
                                                    dataRow[j] = cell.StringCellValue;
                                                    break;
                                            }
                                        }
                                    }
                                    dataTable.Rows.Add(dataRow);
                                }
                            }
                        }
                    }
                }
                return dataTable;
            }
            catch (Exception)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                return null;
            }
        }

        /// <summary>
        /// 将dt中的数据转化为question模型
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>

        /// <summary>
        /// 将DataTable转成Model
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static List<Question> DatatableConvertToQuestion(DataTable dt)
        {
            // 定义集合  
            List<Question> ts = new List<Question>();
            // 获得此模型的类型  
            Type type = typeof(Question);
            string tempName = "";
            foreach (DataRow dr in dt.Rows)
            {
                Question t = new Question();
                // 获得此模型的公共属性  
                PropertyInfo[] propertys = t.GetType().GetProperties();
                foreach (PropertyInfo pi in propertys)
                {
                    tempName = pi.Name;
                    // 检查DataTable是否包含此列  
                    if (dt.Columns.Contains(tempName))
                    {
                        // 判断此属性是否有Setter  
                        if (!pi.CanWrite)
                            continue;
                        object value = dr[tempName];
                        if (value != DBNull.Value)
                        {
                            //pi.SetValue(t, value, null);  
                            pi.SetValue(t, Convert.ChangeType(value, pi.PropertyType, CultureInfo.CurrentCulture), null);
                        }
                    }
                }
                ts.Add(t);
            }
            return ts;
        }


        /// <summary>
        /// 将question对象写入到EXCEL中
        /// </summary>
        /// <param name="question"></param>
        /// <param name="path"></param>
        /// <returns>是否成功</returns>
        public static bool PutQuestionToExcel(Question question, string path)
        {
            FileStream filestream = new FileStream(path, FileMode.Append);
            return PutQuestionToExcel(question, filestream);
        }

        /// <summary>
        /// 将question对象写入到EXCEL中
        /// </summary>
        /// <param name="question"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static bool PutQuestionToExcel(Question question, Stream stream)
        {
            bool mark = false;
            IWorkbook workbook = new HSSFWorkbook();//创建Workbook对象  
            ISheet sheet = workbook.CreateSheet("Sheet1");//创建工作表  
            IRow headerRow = sheet.CreateRow(0);//在工作表中添加首行  
            string[] headerRowName = new string[] { "rownumber", "ID", "SN", "章", "节", "试题", "选项A", "选项B", "选项C", "选项D", "答案", "解析", "备注" };
            ICellStyle style = workbook.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;//设置单元格的样式：水平对齐居中
            IFont font = workbook.CreateFont();//新建一个字体样式对象
                                               // font.Boldweight = short.MaxValue;//设置字体加粗样式
            style.SetFont(font);//使用SetFont方法将字体样式添加到单元格样式中
            for (int i = 0; i < headerRowName.Length; i++)
            {
                ICell cell = headerRow.CreateCell(i);
                cell.SetCellValue(headerRowName[i]);
                cell.CellStyle = style;
            }
            int rownumber = sheet.LastRowNum;
            IRow datarow = sheet.CreateRow(rownumber + 1);
            datarow.CreateCell(0).SetCellValue(rownumber + 1);
            datarow.CreateCell(1).SetCellValue(question.Id);
            datarow.CreateCell(2).SetCellValue(question.SN);
            datarow.CreateCell(3).SetCellValue(new Regex("[_]{3,10}").Replace(question.Chapter, "_______").Trim());
            datarow.CreateCell(4).SetCellValue(question.Node.Trim());
            datarow.CreateCell(5).SetCellValue(question.Title.Trim());
            datarow.CreateCell(6).SetCellValue(question.Choosea.Trim());
            datarow.CreateCell(7).SetCellValue(question.Chooseb.Trim());
            datarow.CreateCell(8).SetCellValue(question.Choosec.Trim());
            datarow.CreateCell(9).SetCellValue(question.Choosed.Trim());
            datarow.CreateCell(10).SetCellValue(question.Answer.Trim());
            datarow.CreateCell(11).SetCellValue(question.Explain.Trim());
            datarow.CreateCell(12).SetCellValue(question.Remark.Trim());
            for (int i = 0; i < headerRow.Cells.Count; i++)
            {
                sheet.AutoSizeColumn(i);
            }
            using (stream)
            {
                workbook.Write(stream);
                stream.Flush();
                stream.Close();
                mark = true;
            }
            return mark;

        }
    }
}

