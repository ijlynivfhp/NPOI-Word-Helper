using Newtonsoft.Json;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using System.Web.Routing;

namespace NPOItest.Models.Sevices
{
    /// <summary>
    /// NpoiHeplper
    /// </summary>
    public class NpoiHeplper
    {
        /// <summary>
        /// 输出模板docx文档(使用字典)
        /// </summary>
        /// <param name="tempFilePath">docx文件路径</param>
        /// <param name="outPath">输出文件路径</param>
        /// <param name="mainData">动态字典数据源</param>
        public static void ExportByDictData(string tempFilePath, string outPath, IDictionary<string, object> mainData, IList<IList<ExpandoObject>> multSubDataList)
        {
            using (FileStream stream = File.OpenRead(tempFilePath))
            {
                XWPFDocument doc = new XWPFDocument(stream);
                //遍历段落  
                if (mainData != null && mainData.Keys.Count > 0)
                    ReplaceParaKey(doc.Paragraphs, mainData);
                //遍历表格
                var docTables = doc.Tables;
                if (docTables.Count > multSubDataList.Count)
                    throw new Exception("Word模板中表格数量与传入数据集合数量不批配！");
                foreach (var docTable in docTables)
                {
                    var tableIndex = docTables.IndexOf(docTable);
                    var subDataList = multSubDataList[tableIndex];
                    var docRow = docTable.GetRow(1);
                    docTable.RemoveRow(1);
                    foreach (var rowData in subDataList)
                    {
                        var addRow = docTable.CreateRow();
                        ReplaceRowKey(docRow, addRow, rowData);
                    }
                }
                //写文件
                FileStream outFile = new FileStream(outPath, FileMode.Create);
                doc.Write(outFile);
                outFile.Close();
            }
        }

        #region 私有方法(动态字典数据源)
        private static void ReplaceRowKey(XWPFTableRow docRow, XWPFTableRow addRow, IDictionary<string, object> dataDict)
        {
            var docCells = docRow.GetTableCells();
            foreach (var docCell in docCells)
            {
                int docCellIndex = docCells.IndexOf(docCell);
                var docCellText = docCell.GetText();
                if (string.IsNullOrEmpty(docCellText))
                    continue;
                foreach (var key in dataDict.Keys)
                {
                    //$$模板中数据占位符为$KEY$
                    if (docCellText.Contains($"${key}$"))
                    {
                        addRow.GetCell(docCellIndex).SetText(docCellText.Replace($"${key}$", dataDict[key]?.ToString()));
                    }
                }
            }
        }
        private static void ReplaceParaKey(IList<XWPFParagraph> paraList, IDictionary<string, object> dataDict)
        {
            foreach (var item in paraList)
            {
                if (item.IsEmpty)
                    continue;
                foreach (var key in dataDict.Keys)
                {
                    //$$模板中数据占位符为$KEY$
                    if (item.ParagraphText.Contains($"${key}$"))
                    {
                        item.ReplaceText($"${key}$", dataDict[key]?.ToString());
                    }
                }
            }
        }
        private static void ReplaceKey(XWPFParagraph para, IDictionary<string, string> data)
        {
            string text = "";
            foreach (var run in para.Runs)
            {
                text = run.ToString();
                foreach (var key in data.Keys)
                {
                    //$$模板中数据占位符为$KEY$
                    if (text.Contains($"${key}$"))
                    {
                        text = text.Replace($"${key}$", data[key]);
                    }
                }
                run.SetText(text, 0);
            }
        }
        #endregion

        /// <summary>
        /// 输出模板docx文档(使用反射)
        /// </summary>
        /// <param name="tempFilePath">docx文件路径</param>
        /// <param name="outPath">输出文件路径</param>
        /// <param name="data">对象数据源</param>
        public static void ExportByObjData(string tempFilePath, string outPath, object mainData, IList<IList<object>> multSubDataList)
        {
            using (FileStream stream = File.OpenRead(tempFilePath))
            {
                XWPFDocument doc = new XWPFDocument(stream);
                //遍历段落      
                if (mainData != null)
                    ReplaceParaObjet(doc.Paragraphs, mainData);
                //遍历表格  
                var docTables = doc.Tables;
                if (docTables.Count > multSubDataList.Count)
                    throw new Exception("Word模板中表格数量与传入数据集合数量不批配！");
                foreach (var docTable in docTables)
                {
                    var tableIndex = docTables.IndexOf(docTable);
                    var subDataList = multSubDataList[tableIndex];
                    var docRow = docTable.GetRow(1);
                    docTable.RemoveRow(1);
                    foreach (var rowData in subDataList)
                    {
                        var addRow = docTable.CreateRow();
                        ReplaceRowObjet(docRow, addRow, rowData);
                    }
                }
                //写文件
                FileStream outFile = new FileStream(outPath, FileMode.Create);
                doc.Write(outFile);
                outFile.Close();
            }
        }

        #region 私有方法（object数据源）
        private static void ReplaceRowObjet(XWPFTableRow docRow, XWPFTableRow addRow, object model)
        {
            Type t = model.GetType();
            PropertyInfo[] pi = t.GetProperties();
            var docCells = docRow.GetTableCells();
            foreach (var docCell in docCells)
            {
                int docCellIndex = docCells.IndexOf(docCell);
                var docCellText = docCell.GetText();
                if (string.IsNullOrEmpty(docCellText))
                    continue;
                foreach (PropertyInfo p in pi)
                {
                    //$$模板中数据占位符为$KEY$
                    string key = $"${p.Name}$";
                    if (docCellText.Contains(key))
                    {
                        try
                        {
                            addRow.GetCell(docCellIndex).SetText(docCellText.Replace(key, p.GetValue(model, null)?.ToString()));
                        }
                        catch (Exception)
                        {
                            //可能有空指针异常
                            continue;
                        }
                    }
                }
            }
        }
        private static void ReplaceParaObjet(IList<XWPFParagraph> paraList, object model)
        {
            foreach (var item in paraList)
            {
                if (item.IsEmpty)
                    continue;
                Type t = model.GetType();
                PropertyInfo[] pi = t.GetProperties();
                foreach (PropertyInfo p in pi)
                {
                    //$$模板中数据占位符为$KEY$
                    string key = $"${p.Name}$";
                    //$$模板中数据占位符为$KEY$
                    if (item.ParagraphText.Contains(key))
                    {
                        try
                        {
                            item.ReplaceText(key, p.GetValue(model, null)?.ToString());
                        }
                        catch (Exception)
                        {
                            //可能有空指针异常
                            continue;
                        }
                    }
                }
            }
        }
        private static void ReplaceKeyObjet(XWPFParagraph para, object model)
        {
            string text = "";
            Type t = model.GetType();
            PropertyInfo[] pi = t.GetProperties();
            foreach (var run in para.Runs)
            {
                text = run.ToString();
                foreach (PropertyInfo p in pi)
                {
                    //$$模板中数据占位符为$KEY$
                    string key = $"${p.Name}$";
                    if (text.Contains(key))
                    {
                        try
                        {
                            text = text.Replace(key, p.GetValue(model, null)?.ToString());
                        }
                        catch (Exception)
                        {
                            //可能有空指针异常
                            text = text.Replace(key, "");
                        }
                    }
                }
                run.SetText(text, 0);
            }
        }
        #endregion


    }
    public static class Extensions
    {
        public static ExpandoObject ToExpando(this object anonymousObject)
        {
            IDictionary<string, object> anonymousDictionary = new RouteValueDictionary(anonymousObject);
            IDictionary<string, object> expando = new ExpandoObject();
            foreach (var item in anonymousDictionary)
                expando.Add(item);
            return (ExpandoObject)expando;
        }
    }
}