using Lucene.Net.Analysis;
using Lucene.Net.Documents;
using Lucene.Net.Index;
using Lucene.Net.QueryParsers;
using Lucene.Net.Search;
using Lucene.Net.Store;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.Text.RegularExpressions;

namespace AI_BAIDU
{
    class LuceneFORnet
    {
        private static FSDirectory directory = FSDirectory.Open(ConfigurationManager.AppSettings["IndexPath"]);
        //private static IndexReader reader = IndexReader.Open(directory, true);
        private static IndexWriter writer = new IndexWriter(directory, new PanGuAnalyzer(), true, IndexWriter.MaxFieldLength.UNLIMITED);
        //private static IndexSearcher searcher = new IndexSearcher(reader);

        #region 分词测试
        /// <summary>
        /// 分词测试
        /// </summary>
        /// <param name="keyword"></param>
        /// <returns></returns>
        public static string Token(string keyword)
        {
            string ret = "";
            System.IO.StringReader reader = new System.IO.StringReader(keyword);
            PanGuAnalyzer analyzer = new PanGuAnalyzer();
            Lucene.Net.Analysis.TokenStream ts = analyzer.TokenStream(keyword, reader);
            bool hasNext = ts.IncrementToken();
            Lucene.Net.Analysis.Tokenattributes.ITermAttribute ita;
            while (hasNext)
            {
                ita = ts.GetAttribute<Lucene.Net.Analysis.Tokenattributes.ITermAttribute>();
                ret += ita.Term + "|";
                hasNext = ts.IncrementToken();
            }
            ts.CloneAttributes();
            reader.Close();
            analyzer.Close();
            return ret;
        }
        #endregion

        public static List<Document> searchIndex(string keyword, int resultCount)
        {
            List<Document> list = null;
            IndexSearcher indexSearcher = new IndexSearcher(directory);
            QueryParser queryParser = new QueryParser(Lucene.Net.Util.Version.LUCENE_30, "title", new PanGuAnalyzer());
            Query query = queryParser.Parse(keyword);
            TopDocs topDocs = indexSearcher.Search(query, resultCount);
            foreach (ScoreDoc doc in topDocs.ScoreDocs)
            {
                list.Add(indexSearcher.Doc(doc.Doc));
            }
            return list;
        }
        /// <summary>
        /// 更新索引，通过先删除，后新增的方式完成
        /// </summary>
        /// <param name="q"></param>
        /// <param name="deleteCount">删除的数量</param>
        /// <returns></returns>
        public static Document updateIndex(Question q, out int deleteCount)
        {
            IndexReader reader = IndexReader.Open(directory, true);
            Term term = new Term("id", q.Subject + q.SNID);
            deleteCount = reader.DeleteDocuments(term);
            Document doc = CreatDOC(q);
            writer.AddDocument(doc);
            writer.Optimize();
            return doc;
        }
        /// <summary>
        /// 生成单个的DOCUMENT
        /// </summary>
        /// <param name="q"></param>
        /// <param name="writer"></param>
        public static Document WriterIndex(Question q)
        {
            IndexWriter writer = LuceneFORnet.writer;
            try
            {
                Document doc = CreatDOC(q);
                writer.AddDocument(doc);
                Console.WriteLine("写入索引:{0}", doc.GetField("id").StringValue);
                return doc;
            }
            catch (FileNotFoundException fnfe)
            {
                throw fnfe;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                // writer.Dispose();
            }
        }
        /// <summary>
        /// 批量生成DOCUMENT
        /// </summary>
        /// <param name="list"></param>
        /// <param name="writer"></param>
        /// <returns></returns>
        public static List<Document> WriterIndex(List<Question> list)
        {
            IndexWriter writer = LuceneFORnet.writer;
            List<Document> listDOC = new List<Document>();
            try
            {
                foreach (Question q in list)
                {
                    Document doc = CreatDOC(q);
                    writer.AddDocument(doc);
                    Console.WriteLine("写入索引:{0}", doc.GetField("id").StringValue);
                    listDOC.Add(doc);
                }
                return listDOC;
            }
            catch (FileNotFoundException fnfe)
            {
                throw fnfe;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //writer.Dispose();
            }
        }
        private static Document CreatDOC(Question q)
        {
            Document doc = new Document();
            string title = new Regex("[_]{5,}").Replace(q.Title, util.GetAnswerStr(q, q.Answer));
            doc.Add(new Field("id", q.Subject + q.SNID, Field.Store.YES, Field.Index.NOT_ANALYZED));//存储且索引
            doc.Add(new Field("title", title, Field.Store.YES, Field.Index.ANALYZED));//存储且索引
            return doc;
        }

        public static void SearchIndex(IndexSearcher searcher)
        {
        }
    }
}
