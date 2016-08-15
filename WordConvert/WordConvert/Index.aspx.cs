using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Data.SQLite;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using FyDB;
using iTextSharp.text.pdf;
using Microsoft.Office.Interop.Word;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using DataTable = System.Data.DataTable;
using Document = iTextSharp.text.Document;
using Paragraph = iTextSharp.text.Paragraph;
using WordConvert.Models;
using iTextSharp.text;

namespace WordConvert
{
    public partial class Index : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Directory.Exists(Server.MapPath("~/wordInfo/")))
                Directory.CreateDirectory(Server.MapPath("~/wordInfo/"));

            if (!Directory.Exists(Server.MapPath("~/log/")))
                Directory.CreateDirectory(Server.MapPath("~/log/"));
        }

        //转换成html按钮
        protected void Button2_Click(object sender, EventArgs e)
        {
            //创建临时文件，避开浏览器不兼容问题
            string fname = Server.MapPath("~/wordInfo/") + Guid.NewGuid().ToString() + ".doc";
            this.FileUpload2.SaveAs(fname);

            GetPathByDocToHTML(fname);
        }

        #region wordFormToHtml

        /// <summary>
        /// word转换html
        /// </summary>
        /// <param name="strFile">全路径</param>
        /// <returns></returns>
        private string GetPathByDocToHTML(string strFile)
        {
            try
            {
                if (string.IsNullOrEmpty(strFile))
                {
                    return "0"; //没有文件
                }

                Microsoft.Office.Interop.Word.ApplicationClass word = new Microsoft.Office.Interop.Word.ApplicationClass();
                Type wordType = word.GetType();
                Microsoft.Office.Interop.Word.Documents docs = word.Documents;

                // 打开文件  
                Type docsType = docs.GetType();

                object fileName = strFile;

                Microsoft.Office.Interop.Word.Document doc =
                    (Microsoft.Office.Interop.Word.Document)docsType.InvokeMember("Open",
                        System.Reflection.BindingFlags.InvokeMethod, null, docs, new Object[] { fileName, true, true });

                // 转换格式，另存为html  
                Type docType = doc.GetType();
                //给文件重新起名
                string filename = System.DateTime.Now.Year.ToString() + System.DateTime.Now.Month.ToString() +
                                  System.DateTime.Now.Day.ToString() +
                                  System.DateTime.Now.Hour.ToString() + System.DateTime.Now.Minute.ToString() +
                                  System.DateTime.Now.Second.ToString();

                string strFileFolder = "~/html/";
                DateTime dt = DateTime.Now;
                //以yyyymmdd形式生成子文件夹名
                string strFileSubFolder = dt.Year.ToString();
                strFileSubFolder += (dt.Month < 10) ? ("0" + dt.Month.ToString()) : dt.Month.ToString();
                strFileSubFolder += (dt.Day < 10) ? ("0" + dt.Day.ToString()) : dt.Day.ToString();
                string strFilePath = strFileFolder + strFileSubFolder + "/";
                // 判断指定目录下是否存在文件夹，如果不存在，则创建 
                if (!Directory.Exists(Server.MapPath(strFilePath)))
                {
                    // 创建up文件夹 
                    Directory.CreateDirectory(Server.MapPath(strFilePath));
                }

                //被转换的html文档保存的位置 
                // HttpContext.Current.Server.MapPath("html" + strFileSubFolder + filename + ".html")
                string ConfigPath = Server.MapPath(strFilePath + filename + ".html");
                object saveFileName = ConfigPath;

                /*下面是Microsoft Word 9 Object Library的写法，如果是10，可能写成： 
                  * docType.InvokeMember("SaveAs", System.Reflection.BindingFlags.InvokeMethod, 
                  * null, doc, new object[]{saveFileName, Word.WdSaveFormat.wdFormatFilteredHTML}); 
                  * 其它格式： 
                  * wdFormatHTML 
                  * wdFormatDocument 
                  * wdFormatDOSText 
                  * wdFormatDOSTextLineBreaks 
                  * wdFormatEncodedText 
                  * wdFormatRTF 
                  * wdFormatTemplate 
                  * wdFormatText 
                  * wdFormatTextLineBreaks 
                  * wdFormatUnicodeText 
                */
                docType.InvokeMember("SaveAs", System.Reflection.BindingFlags.InvokeMethod,
                    null, doc, new object[] { saveFileName, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatFilteredHTML });

                //docType.InvokeMember("SaveAs", System.Reflection.BindingFlags.InvokeMethod,
                //  null, doc, new object[] { saveFileName, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatFilteredHTML }); 

                //关闭文档  
                docType.InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod,
                    null, doc, new object[] { null, null, null });

                // 退出 Word  
                wordType.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, word, null);
                //转到新生成的页面  
                //return ("/" + filename + ".html");

                //转化HTML页面统一编码格式
                TransHTMLEncoding(ConfigPath);

                File.Delete(strFile); //删除临时保存的文件

                return (strFilePath + filename + ".html");
            }
            catch (Exception ex)
            {
                using (FileStream fs = new FileStream(Server.MapPath("~/log/") + DateTime.Now.ToString("yyyyMMdd") + ".txt", FileMode.OpenOrCreate, FileAccess.Write))
                {
                    using (StreamWriter sw = new StreamWriter(fs))
                    {
                        sw.WriteLine(ex.ToString()+"\n"+ex.Message.ToString()+"\n"+ex.StackTrace+"\n"+ex.InnerException+"\n");
                    };
                };

                File.Delete(strFile); //删除临时保存的文件
                Page.RegisterStartupScript("alt", "<script>alert('转Html出错了。')</script>");
                return "0"; //没有文件
            }
        }

        /// <summary>
        /// 编码转换
        /// </summary>
        /// <param name="strFilePath"></param>
        private void TransHTMLEncoding(string strFilePath)
        {
            try
            {
                System.IO.StreamReader sr = new System.IO.StreamReader(strFilePath, Encoding.GetEncoding(0));
                string html = sr.ReadToEnd();
                sr.Close();
                html = System.Text.RegularExpressions.Regex.Replace(html, @"<meta[^>]*>",
                    "<meta http-equiv=Content-Type content='text/html; charset=gb2312'>",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                System.IO.StreamWriter sw = new System.IO.StreamWriter(strFilePath, false, Encoding.Default);

                sw.Write(html);
                sw.Close();
            }
            catch (Exception ex)
            {
                Page.RegisterStartupScript("alt", "<script>alert('" + ex.Message + "')</script>");
            }
        }

        #endregion

        #region DataTableFormToExcel

        /// <summary>
        /// 转换为excel
        /// </summary>
        /// <param name="tbName"></param>
        public void FormToExcel(string tbName, DataTable tblDatas)
        {
            NPOI.HSSF.UserModel.HSSFWorkbook book = new NPOI.HSSF.UserModel.HSSFWorkbook();
            var sheetReportResult = book.CreateSheet(tbName);
            
            sheetReportResult.CreateRow(0).CreateCell(0).SetCellValue("姓名");
            sheetReportResult.GetRow(0).CreateCell(1).SetCellValue("性别");
            sheetReportResult.GetRow(0).CreateCell(2).SetCellValue("年龄");
            sheetReportResult.GetRow(0).CreateCell(3).SetCellValue("学历");
            sheetReportResult.GetRow(0).CreateCell(4).SetCellValue("手机");
            sheetReportResult.GetRow(0).CreateCell(5).SetCellValue("电子邮件");
            sheetReportResult.GetRow(0).CreateCell(6).SetCellValue("英语等级");
            sheetReportResult.GetRow(0).CreateCell(7).SetCellValue("求职意向");
            sheetReportResult.GetRow(0).CreateCell(8).SetCellValue("工作地点");
            sheetReportResult.GetRow(0).CreateCell(9).SetCellValue("工作年限");
            sheetReportResult.GetRow(0).CreateCell(10).SetCellValue("期望薪水");
            sheetReportResult.GetRow(0).CreateCell(11).SetCellValue("毕业院校");
            sheetReportResult.GetRow(0).CreateCell(12).SetCellValue("专业");
            sheetReportResult.GetRow(0).CreateCell(13).SetCellValue("最近单位");
            sheetReportResult.GetRow(0).CreateCell(14).SetCellValue("最近职位");
            sheetReportResult.GetRow(0).CreateCell(15).SetCellValue("接收时间");

            //循环内容
            for (int i = 0; i < tblDatas.Rows.Count; i++)
            {
                sheetReportResult.CreateRow(i + 1).CreateCell(0).SetCellValue(tblDatas.Rows[i][0].ToString());
                for (int j = 0; j < tblDatas.Columns.Count; j++)
                {
                    sheetReportResult.GetRow(i + 1).CreateCell(j).SetCellValue(tblDatas.Rows[i][j].ToString());
                }
            }

            sheetReportResult.SetColumnWidth(0, 20 * 256);
            sheetReportResult.SetColumnWidth(1, 40 * 256);
            sheetReportResult.SetColumnWidth(2, 40 * 256);

            // 写入到客户端  
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            book.Write(ms);
            Response.AddHeader("Content-Disposition",
                string.Format("attachment; filename={0}.xls", DateTime.Now.ToString("yyyyMMddHHmmssfff")));
            Response.BinaryWrite(ms.ToArray());
            book = null;
            ms.Close();
            ms.Dispose();
        }

        #endregion

        #region DataTableFormToWord

        public void ExportDataGridViewToWord(DataTable srcDgv)
        {
            if (srcDgv.Rows.Count == 0)
            {
                Page.RegisterStartupScript("alt", "<script>alert('没有数据可供导出!')</script>");
                return;
            }
            else
            {
                Object none = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document document = wordApp.Documents.Add(ref none, ref none, ref none, ref none);
                //建立表格
                Microsoft.Office.Interop.Word.Table table = document.Tables.Add(document.Paragraphs.Last.Range, srcDgv.Rows.Count + 1, srcDgv.Columns.Count, ref none, ref none);

                try
                {
                    table.Cell(1, 1).Range.Text = "姓名";
                    table.Cell(1, 2).Range.Text = "性别";
                    table.Cell(1, 3).Range.Text = "年龄";
                    table.Cell(1, 4).Range.Text = "学历";
                    table.Cell(1, 5).Range.Text = "手机";
                    table.Cell(1, 6).Range.Text = "电子邮件";
                    table.Cell(1, 7).Range.Text = "英语等级";
                    table.Cell(1, 8).Range.Text = "求职意向";
                    table.Cell(1, 9).Range.Text = "工作地点";
                    table.Cell(1, 10).Range.Text = "工作年限";
                    table.Cell(1, 11).Range.Text = "期望薪水";
                    table.Cell(1, 12).Range.Text = "毕业院校";
                    table.Cell(1, 13).Range.Text = "专业";
                    table.Cell(1, 14).Range.Text = "最近单位";
                    table.Cell(1, 15).Range.Text = "最近职位";
                    table.Cell(1, 16).Range.Text = "接收时间";
                    //输出控件中的记录
                    for (int i = 0; i < srcDgv.Rows.Count; i++)
                    {
                        for (int j = 0; j < srcDgv.Columns.Count; j++)
                        {
                            table.Cell(i + 2, j + 1).Range.Text = srcDgv.Rows[i][j].ToString();
                        }
                    }
                    table.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                    table.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                    string newFile = Server.MapPath("~/wordInfo/") + DateTime.Now.ToString("yyyyMMddHHmmssss") + ".doc";
                    document.SaveAs(newFile, ref none, ref none, ref none, ref none, ref none, ref none, ref none, ref none, ref none, ref none, ref none, ref none, ref none, ref none, ref none);

                    Page.RegisterStartupScript("alt", "<script>alert('数据成功导出!')</script>");
                }
                catch (Exception e)
                {
                    Page.RegisterStartupScript("alt", "<script>alert('" + e.Message + "')</script>");
                }
            }
        }
        #endregion

        #region DataTableFormToPdf

        public void FormToPdf(DataTable srcDgv)
        {
            //iTextSharp.text.Rectangle pageSize = new iTextSharp.text.Rectangle(144, 720);
            iTextSharp.text.Rectangle pageSize = new iTextSharp.text.Rectangle(1366, 720);
            pageSize.BackgroundColor = new Color(0xFF, 0xFF, 0xDE);

            Document document = new Document(pageSize);

            string newFile = Server.MapPath("~/wordInfo/") + DateTime.Now.ToString("yyyyMMddHHmmssss") + ".pdf";
            //string newFile = "c://" + DateTime.Now.ToString("yyyyMMddHHmmssss") + ".pdf";
            PdfWriter.GetInstance(document, new FileStream(newFile, FileMode.Create));
            document.Open();
            BaseFont bfChinese = BaseFont.CreateFont("C://WINDOWS//Fonts//simsun.ttc,1", BaseFont.IDENTITY_H,
                BaseFont.NOT_EMBEDDED);
            iTextSharp.text.Font fontChinese = new iTextSharp.text.Font(bfChinese, 12,
                iTextSharp.text.Font.NORMAL, new iTextSharp.text.Color(0, 0, 0));

            StringBuilder sbBuilder = new StringBuilder();
            
            sbBuilder.Append("姓名   ");
            sbBuilder.Append("性别   ");
            sbBuilder.Append("年龄   ");
            sbBuilder.Append("学历   ");
            sbBuilder.Append("手机   ");
            sbBuilder.Append("电子邮件   ");
            sbBuilder.Append("英语等级   ");
            sbBuilder.Append("求职意向   ");
            sbBuilder.Append("工作地点   ");
            sbBuilder.Append("工作年限   ");
            sbBuilder.Append("期望薪水   ");
            sbBuilder.Append("毕业院校   ");
            sbBuilder.Append("专业   ");
            sbBuilder.Append("最近单位   ");
            sbBuilder.Append("最近职位   ");            
            sbBuilder.Append("接收时间   ");
            sbBuilder.Append("\n");

            //输出控件中的记录
            for (int i = 0; i < srcDgv.Rows.Count; i++)
            {
                for (int j = 0; j < srcDgv.Columns.Count; j++)
                {
                    sbBuilder.Append(srcDgv.Rows[i][j].ToString().Trim() + "   ");
                }
                sbBuilder.Append("\n");
            }
            //导出文本的内容：
            document.Add(new Paragraph(sbBuilder.ToString(), fontChinese));
            //导出图片：
            //iTextSharp.text.Image jpeg = iTextSharp.text.Image.GetInstance(Path.GetFullPath("1.jpg"));
            //document.Add(jpeg);

            //注意一定要关闭，否则PDF中的内容将得不到保存

            document.Close();
            Page.RegisterStartupScript("alt", "<script>alert('数据成功导出!')</script>");
        }
        #endregion

        public static DataTable ConvertDataReaderToDataTable(SqlDataReader reader)
        {
            try
            {
                DataTable objDataTable = new DataTable();
                int intFieldCount = reader.FieldCount;
                for (int intCounter = 0; intCounter < intFieldCount; ++intCounter)
                {
                    objDataTable.Columns.Add(reader.GetName(intCounter), reader.GetFieldType(intCounter));
                }
                objDataTable.BeginLoadData();

                object[] objValues = new object[intFieldCount];
                while (reader.Read())
                {
                    reader.GetValues(objValues);
                    objDataTable.LoadDataRow(objValues, true);
                }
                reader.Close();
                objDataTable.EndLoadData();

                return objDataTable;

            }
            catch (Exception ex)
            {
                throw new Exception("转换出错!", ex);
            }

        }

        //导出按钮
        protected void Button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.tbName.Text != "")
                {
                    string sql = string.Format("select * from {0}", this.tbName.Text);
                    SqlDataReader reader = SqlHelper.ExecuteReader(SqlHelper.ConnectionStringLocalTransaction, CommandType.Text,
                             sql, null);
                    DataTable tblDatas = ConvertDataReaderToDataTable(reader);
                    switch (this.DropDownList1.SelectedIndex)
                    {
                        case 0: //word
                            ExportDataGridViewToWord(tblDatas);
                            break;
                        case 1: //excel
                            FormToExcel(this.tbName.Text, tblDatas);
                            break;
                        case 2: //pdf
                            FormToPdf(tblDatas);
                            break;
                    }
                }
                else
                {
                    Page.RegisterStartupScript("alt", "<script>alert('请填写要转换的表名!')</script>");
                }
            }
            catch (Exception ex)
            {
                Page.RegisterStartupScript("alt", "<script>alert('转换出错了!')</script>");
            }
        }

        //提取word内容
        protected void Button1_Click(object sender, EventArgs e)
        {
            //创建临时文档
            string fname = Server.MapPath("~/wordInfo/") + Guid.NewGuid().ToString() + ".doc";
            try
            {                
                //创建临时文档
                this.FileUpload1.SaveAs(fname);

                //创建word
                _Application app = new Microsoft.Office.Interop.Word.Application();
                //创建word文档
                _Document doc = null;
                object unknow = Type.Missing;
                doc = app.Documents.Open(fname,
                               ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                               ref unknow, ref unknow, ref unknow, ref unknow, ref unknow,
                               ref unknow, ref unknow, ref unknow, ref unknow, ref unknow);
                string content = doc.Content.Text;

                //去除文档中的回车、换行、制表符
                string clearContent = content
                    .Replace("\n", string.Empty)
                    .Replace("\r", string.Empty)
                    .Replace("\a", string.Empty)
                    .Replace("\t", string.Empty);
                //.Replace(" ",string.Empty);

                CurriculumVitae cv = new CurriculumVitae();
                cv.ReceiptTime = DateTime.Now;

                if (this.DropDownList2.SelectedItem.Text == "中国皮革人才网")
                {
                    #region 皮革人才网            
                    string subContent = clearContent.Substring(clearContent.IndexOf("标准简历"));
                    string subContentTitle = clearContent.Substring(clearContent.IndexOf("简历来源"), clearContent.IndexOf("标准简历"));

                    //工作年限
                    cv.WorkLife = subContentTitle.Substring(subContentTitle.IndexOf("元/月")
                        , subContentTitle.IndexOf("年") - subContentTitle.IndexOf("元/月"))
                        .Substring("元/月".Length, subContentTitle.Substring(subContentTitle.IndexOf("元/月")
                        , subContentTitle.IndexOf("年") - subContentTitle.IndexOf("元/月")).Length - "元/月".Length);

                    //姓名
                    cv.Name = subContent.Substring(subContent.IndexOf("姓　　名：")
                        , subContent.IndexOf("性　　别：") - subContent.IndexOf("姓　　名：") - 1).Split('：')[1];

                    //性别
                    cv.Sex = subContent.Substring(subContent.IndexOf("性　　别：")
                        , subContent.IndexOf("婚姻状况： ") - subContent.IndexOf("性　　别：") - 1).Split('：')[1];

                    //年龄
                    cv.Age = subContent.Substring(subContent.IndexOf("年　　龄：")
                        , subContent.IndexOf("现所在地：") - subContent.IndexOf("年　　龄：") - 1).Split('：')[1];

                    //学历
                    cv.Education = subContent.Contains("初中") ? "初中以下" :
                        subContent.Contains("高中") ? "高中" :
                        subContent.Contains("中专") ? "中专" :
                        subContent.Contains("大专") ? "大专" :
                        subContent.Contains("本科") ? "本科" :
                        subContent.Contains("硕士") ? "硕士" : "博士";

                    //手机号码
                    cv.Phone = subContent.Substring(subContent.IndexOf("手机号码： ")
                        , subContent.IndexOf("电子邮件： ") - subContent.IndexOf("手机号码： ") - 1).Split('：')[1];

                    //电子邮件
                    cv.Email = subContent.Substring(subContent.IndexOf("电子邮件： ")
                        , subContent.IndexOf("后备电话： ") - subContent.IndexOf("电子邮件： ") - 1).Split('：')[1];

                    //英语等级
                    cv.EnglishLevel = subContent.Substring(subContent.IndexOf("英语水平：")
                        , subContent.IndexOf("英语：   ") - subContent.IndexOf("英语水平：") - 1).Split('：')[1];

                    //求职意向
                    cv.JobIntension = subContent.Substring(subContent.IndexOf("应聘职位： ")
                        , subContent.IndexOf("意向职位：") - subContent.IndexOf("应聘职位： ")).Split('：')[1];

                    //期望薪水
                    cv.Salary = subContent.Substring(subContent.IndexOf("待遇要求：")
                        , subContent.IndexOf("求职状态：") - subContent.IndexOf("待遇要求：") - 1).Split('：')[1];

                    //毕业院校
                    cv.School = subContent.Substring(subContent.IndexOf("教育/培训"),
                        subContent.IndexOf("语言能力") - subContent.IndexOf("教育/培训") - 1).Split(' ')[1];

                    //专业
                    cv.Major = subContent.Substring(subContent.IndexOf("教育/培训"),
                        subContent.IndexOf("语言能力") - subContent.IndexOf("教育/培训") - 1).Split(' ')[2];

                    //工作地点
                    cv.WorkPlace = subContent.Substring(subContent.IndexOf("意向地区： ")
                        , subContent.IndexOf("工作性质：") - subContent.IndexOf("意向地区： ") - 1).Split('：')[1];

                    //最近单位
                    cv.RecentWorkUnits = string.Empty;

                    //最近职位
                    cv.RecnetJob = string.Empty;
                    #endregion 皮革人才网
                }
                else if (this.DropDownList2.SelectedItem.Text == "前程无忧")
                {
                    #region 前程无忧            
                    string subContent = clearContent;

                    //工作年限
                    cv.WorkLife = clearContent.Substring(clearContent.IndexOf("%"),
                        clearContent.IndexOf("工作经验 ") - clearContent.IndexOf("%") - 1).Split(' ')[1];

                    //姓名
                    cv.Name = subContent.Substring(subContent.IndexOf("更新时间：")
                        , subContent.IndexOf("匹配度") - subContent.IndexOf("更新时间：") - 1).Split('/')[1];

                    //性别
                    cv.Sex = subContent.Substring(subContent.IndexOf("工作经验 | ") + "工作经验 | ".Length, 1);

                    //年龄
                    cv.Age = subContent.Substring(subContent.IndexOf(cv.Sex + " |  ") + (cv.Sex + " |  ").Length,
                        subContent.IndexOf("岁") - subContent.IndexOf(cv.Sex + " |  ") - (cv.Sex + " |  ").Length);

                    //学历
                    cv.Education = subContent.Substring(subContent.IndexOf("学　历："),
                        subContent.IndexOf("专　业：") - subContent.IndexOf("学　历：")).Split('：')[1];

                    //手机号码
                    cv.Phone = subContent.Substring(subContent.IndexOf("电　话：")
                        , subContent.IndexOf("E-mail：") - subContent.IndexOf("电　话：") - 1).Split('：')[1];

                    //电子邮件
                    cv.Email = subContent.Substring(subContent.IndexOf("E-mail：")
                        , subContent.IndexOf("最近工作") - subContent.IndexOf("E-mail：")).Split('：')[1];

                    //英语等级
                    cv.EnglishLevel = string.Empty;
                    //cv.EnglishLevel = subContent.Substring(subContent.IndexOf("英语水平：")
                    //    , subContent.IndexOf("英语：   ") - subContent.IndexOf("英语水平：") - 1).Split('：')[1];

                    //求职意向
                    cv.JobIntension = subContent.Substring(subContent.IndexOf("目标职能： ")
                        , subContent.IndexOf("求职状态") - subContent.IndexOf("目标职能： ")).Split('：')[1];

                    //期望薪水
                    cv.Salary = subContent.Substring(subContent.IndexOf("期望薪资： ")
                        , subContent.IndexOf("目标职能") - subContent.IndexOf("期望薪资： ")).Split('：')[1];

                    //毕业院校
                    cv.School = subContent.Substring(subContent.IndexOf("学　校："),
                        subContent.IndexOf("自我评价") - subContent.IndexOf("学　校：") - 1).Split('：')[1];

                    //专业
                    cv.Major = subContent.Substring(subContent.IndexOf("专　业："),
                        subContent.IndexOf("学　校") - subContent.IndexOf("专　业：")).Split('：')[1];

                    //工作地点
                    cv.WorkPlace = subContent.Substring(subContent.IndexOf("目标地点： ")
                        , subContent.IndexOf("期望薪资") - subContent.IndexOf("目标地点： ")).Split('：')[1];

                    //最近单位
                    cv.RecentWorkUnits = subContent.Substring(subContent.IndexOf("公　司：")
                        , subContent.IndexOf("行　业") - subContent.IndexOf("公　司：")).Split('：')[1];

                    //最近职位
                    cv.RecnetJob = subContent.Substring(subContent.IndexOf("职　位：")
                        , subContent.IndexOf("学历学　历") - subContent.IndexOf("职　位：")).Split('：')[1];
                    #endregion 前程无忧
                }
                else if (this.DropDownList2.SelectedItem.Text == "猎聘网")
                {
                    #region 猎聘网            
                    string subContent = clearContent;

                    //工作年限
                    cv.WorkLife = subContent.Substring(subContent.IndexOf("工作年限：")
                        , subContent.IndexOf("年婚姻状况") - subContent.IndexOf("工作年限：") - 1).Split('：')[1];

                    //姓名
                    cv.Name = subContent.Substring(subContent.IndexOf("姓名：")
                        , subContent.IndexOf("性别") - subContent.IndexOf("姓名：")).Split('：')[1];

                    //性别
                    cv.Sex = subContent.Substring(subContent.IndexOf("性别：")
                        , subContent.IndexOf("/手机号码") - subContent.IndexOf("性别：")).Split('：')[1];

                    //年龄
                    cv.Age = subContent.Substring(subContent.IndexOf("年龄：")
                        , subContent.IndexOf("岁电子邮件") - subContent.IndexOf("年龄：") - 1).Split('：')[1];

                    //学历
                    cv.Education = subContent.Substring(subContent.IndexOf("教育程度：")
                        , subContent.IndexOf("工作年限") - subContent.IndexOf("教育程度：")).Split('：')[1];

                    //手机号码
                    cv.Phone = subContent.Substring(subContent.IndexOf("手机号码：")
                        , subContent.IndexOf("年龄") - subContent.IndexOf("手机号码：") - 1).Split('：')[1];

                    //电子邮件
                    cv.Email = subContent.Substring(subContent.IndexOf("电子邮件：")
                        , subContent.IndexOf("教育程度") - subContent.IndexOf("电子邮件：")).Split('：')[1];

                    //英语等级
                    cv.EnglishLevel = string.Empty;
                    //cv.EnglishLevel = subContent.Substring(subContent.IndexOf("英语水平：")
                    //    , subContent.IndexOf("英语：   ") - subContent.IndexOf("英语水平：") - 1).Split('：')[1];

                    //求职意向
                    cv.JobIntension = subContent.Substring(subContent.IndexOf("期望职位：")
                        , subContent.IndexOf("期望地点") - subContent.IndexOf("期望职位：")).Split('：')[1];

                    //期望薪水
                    cv.Salary = subContent.Substring(subContent.IndexOf("期望年薪：")
                        , subContent.IndexOf("工作经历") - subContent.IndexOf("期望年薪：")).Split('：')[1];

                    //毕业院校
                    cv.School = subContent.Substring(subContent.IndexOf("教育经历"),
                        subContent.IndexOf("语言能力") - subContent.IndexOf("教育经历")).Split(new char[] { ' ', ' ' })[0]
                        .Replace("教育经历", string.Empty);

                    //专业
                    cv.Major = subContent.Substring(subContent.IndexOf("教育经历"),
                        subContent.IndexOf("语言能力") - subContent.IndexOf("教育经历")).Split(new char[] { ' ', ' ' })[2]
                        .Replace(cv.Education, string.Empty).Substring(15);

                    //工作地点
                    cv.WorkPlace = subContent.Substring(subContent.IndexOf("期望地点：")
                        , subContent.IndexOf("期望年薪：") - subContent.IndexOf("期望地点：")).Split('：')[1];

                    //最近单位
                    cv.RecentWorkUnits = subContent.Substring(subContent.IndexOf("公司名称：")
                        , subContent.IndexOf("所任职位：") - subContent.IndexOf("公司名称：")).Split('：')[1];

                    //最近职位
                    cv.RecnetJob = subContent.Substring(subContent.IndexOf("所任职位：")
                        , subContent.IndexOf("目前年薪") - subContent.IndexOf("所任职位：")).Split('：')[1];
                    #endregion 猎聘网
                }
                else if (this.DropDownList2.SelectedItem.Text == "国际人才")
                {
                    #region 国际人才            
                    string subContent = clearContent;

                    //工作年限
                    cv.WorkLife = subContent.Substring(subContent.IndexOf("工作经验: "), subContent.IndexOf("| 现居住地") - subContent.IndexOf("工作经验: ") - 1).Split(':')[1];

                    //姓名
                    cv.Name = subContent.Substring(subContent.IndexOf("联系人：")
                        , subContent.IndexOf("联系电话") - subContent.IndexOf("联系人：")).Split('：')[1];

                    //性别
                    cv.Sex = subContent.Substring(subContent.IndexOf("性别:")
                        , subContent.IndexOf("| 身高") - subContent.IndexOf("性别:") - 1).Split(':')[1];

                    //年龄
                    cv.Age = subContent.Substring(subContent.IndexOf("年龄： ")
                        , subContent.IndexOf("岁") - subContent.IndexOf("年龄： ")).Split('：')[1];

                    //学历
                    cv.Education = subContent.Substring(subContent.IndexOf("最高学历：")
                        , subContent.IndexOf("| \v工作经验") - subContent.IndexOf("最高学历：")).Split('：')[1];

                    //手机号码
                    cv.Phone = subContent.Substring(subContent.IndexOf("联系电话：")
                        , subContent.IndexOf("联系邮箱") - subContent.IndexOf("联系电话：")).Split('：')[1];

                    //电子邮件
                    cv.Email = subContent.Substring(subContent.IndexOf("联系邮箱：")
                        , subContent.IndexOf("http:") - subContent.IndexOf("联系邮箱：")).Split('：')[1];

                    //英语等级
                    cv.EnglishLevel = string.Empty;
                    //cv.EnglishLevel = subContent.Substring(subContent.IndexOf("英语水平：")
                    //    , subContent.IndexOf("英语：   ") - subContent.IndexOf("英语水平：") - 1).Split('：')[1];

                    //求职意向
                    cv.JobIntension = subContent.Substring(subContent.IndexOf("期望从事的岗位：")
                        , subContent.IndexOf("\v期望从事的行业") - subContent.IndexOf("期望从事的岗位：")).Split('：')[1];

                    //期望薪水
                    cv.Salary = subContent.Substring(subContent.IndexOf("期望月薪：")
                        , subContent.IndexOf("\v期望从事的岗位") - subContent.IndexOf("期望月薪：")).Split('：')[1];

                    //毕业院校
                    cv.School = subContent.Substring(subContent.IndexOf("教育经历"),
                        subContent.IndexOf("专业描述：") - subContent.IndexOf("教育经历")).Split('|')[0].Substring(20);

                    //专业
                    cv.Major = subContent.Substring(subContent.IndexOf("教育经历"),
                        subContent.IndexOf("专业描述：") - subContent.IndexOf("教育经历")).Split('|')[1];

                    //工作地点
                    cv.WorkPlace = string.Empty;
                    //cv.WorkPlace = subContent.Substring(subContent.IndexOf("意向地区： ")
                    //    , subContent.IndexOf("工作性质：") - subContent.IndexOf("意向地区： ") - 1).Split('：')[1];

                    //最近单位
                    cv.RecentWorkUnits = string.Empty;

                    //最近职位
                    cv.RecnetJob = string.Empty;
                    #endregion 国际人才
                }


                doc.Close();

                File.Delete(fname);//删除临时文件

                string sql = @"Insert into TestData 
(Name,Sex,Age,Education,Phone,Email,EnglishLevel,JobIntension,WorkPlace,WorkLife,
Salary,School,Major,RecentWorkUnits,RecnetJob,ReceiptTime) 
values ('" + cv.Name.Trim() + "','" + cv.Sex.Trim() + "','" + cv.Age.Trim() + "','" + cv.Education.Trim() + "','" + cv.Phone.Trim() +
    "','" + cv.Email.Trim() + "','" + cv.EnglishLevel.Trim() + "','" + cv.JobIntension.Trim() + "','" + cv.WorkPlace.Trim() + "','" +
    cv.WorkLife.Trim() + "','" + cv.Salary.Trim() + "','" + cv.School.Trim() + "','" + cv.Major.Trim() + "','" + cv.RecentWorkUnits.Trim() +
    "','" + cv.RecnetJob.Trim() + "','" + cv.ReceiptTime + "')";

                var result = SqlHelper.ExecuteNonQuery(SqlHelper.ConnectionStringLocalTransaction, CommandType.Text, sql, null);

                Page.RegisterStartupScript("alt", "<script>alert('完成：" + result.ToString() + "!')</script>");
            }
            catch (Exception ex)
            {
                using (FileStream fs = new FileStream(Server.MapPath("~/log/") + DateTime.Now.ToString("yyyyMMdd") + ".txt", FileMode.OpenOrCreate, FileAccess.Write))
                {
                    using (StreamWriter sw = new StreamWriter(fs))
                    {
                        sw.WriteLine(ex.ToString()+"\n"+ex.Message.ToString()+"\n"+ex.StackTrace+"\n"+ex.InnerException+"\n");
                    };
                };

                File.Delete(fname); //删除临时保存的文件
                Page.RegisterStartupScript("alt", "<script>alert('提取出错了。')</script>");
            }
        }
    }
}
