using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using iTextSharp.text;
using iTextSharp.text.pdf;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;

namespace WordConvert
{
    public partial class Index : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

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

        /// <summary>
        /// 转换为excel
        /// </summary>
        /// <param name="tbName"></param>
        public void FormToExcel(string tbName)
        {
            //测试连接数据库
            SQLiteConnection conn = null;
            string dbPath = "Data Source =" + Server.MapPath("~/sql/") + "/wordTest.db";
            //dbPath = @"Data Source =E:\jww\扶摇\项目\转换word\sql\wordTest";
            //conn = new SQLiteConnection(dbPath);//创建数据库实例，指定文件位置  
            //conn.Open();//打开数据库，若文件不存在会自动创建  

            //string sql = "select * from WordTable";
            //SQLiteCommand cmdQ = new SQLiteCommand(sql, conn);

            //SQLiteDataReader reader = cmdQ.ExecuteReader();
            //while (reader.Read())
            //{
            //    Console.WriteLine(reader.GetInt32(0) + " " + reader.GetString(1) + " " + reader.GetString(2));
            //}
            //conn.Close();

            //测试假数据
            DataTable tblDatas = new DataTable("Datas");
            DataColumn dc = null;
            dc = tblDatas.Columns.Add("ID", Type.GetType("System.Int32"));
            dc.AutoIncrement = true;//自动增加
            dc.AutoIncrementSeed = 1;//起始为1
            dc.AutoIncrementStep = 1;//步长为1
            dc.AllowDBNull = false;//

            dc = tblDatas.Columns.Add("Product", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("Version", Type.GetType("System.String"));
            dc = tblDatas.Columns.Add("Description", Type.GetType("System.String"));

            DataRow newRow;
            newRow = tblDatas.NewRow();
            newRow["Product"] = "大话西游";
            newRow["Version"] = "2.0";
            newRow["Description"] = "我很喜欢";
            tblDatas.Rows.Add(newRow);

            newRow = tblDatas.NewRow();
            newRow["Product"] = "梦幻西游";
            newRow["Version"] = "3.0";
            newRow["Description"] = "比大话更幼稚";
            tblDatas.Rows.Add(newRow);
            NPOI.HSSF.UserModel.HSSFWorkbook book = new NPOI.HSSF.UserModel.HSSFWorkbook();
            var sheetReportResult = book.CreateSheet(tbName);

            //产生第一个要用CreateRow 
            sheetReportResult.CreateRow(0).CreateCell(0).SetCellValue("ID");
            //之后的用GetRow 取得在CreateCell
            sheetReportResult.GetRow(0).CreateCell(1).SetCellValue("Product");
            sheetReportResult.GetRow(0).CreateCell(2).SetCellValue("Version");
            sheetReportResult.GetRow(0).CreateCell(3).SetCellValue("Description");

            //循环内容
            for (var i = 0; i < tblDatas.Rows.Count; i++)
            {
                sheetReportResult.CreateRow(i).CreateCell(0).SetCellValue(tblDatas.Rows[i]["ID"].ToString());
                sheetReportResult.GetRow(i).CreateCell(1).SetCellValue(tblDatas.Rows[i]["Product"].ToString());
                sheetReportResult.GetRow(i).CreateCell(2).SetCellValue(tblDatas.Rows[i]["Version"].ToString());
                sheetReportResult.GetRow(i).CreateCell(3).SetCellValue(tblDatas.Rows[i]["Description"].ToString());
            }

            sheetReportResult.SetColumnWidth(0, 20 * 256);
            sheetReportResult.SetColumnWidth(1, 40 * 256);
            sheetReportResult.SetColumnWidth(2, 40 * 256);

            // 写入到客户端  
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            book.Write(ms);
            Response.AddHeader("Content-Disposition", string.Format("attachment; filename={0}.xls", DateTime.Now.ToString("yyyyMMddHHmmssfff")));
            Response.BinaryWrite(ms.ToArray());
            book = null;
            ms.Close();
            ms.Dispose();
        }

        //导出按钮
        protected void Button3_Click(object sender, EventArgs e)
        {
            switch (this.DropDownList1.SelectedIndex)
            {
                case 0://word

                    break;
                case 1://excel
                    FormToExcel(this.tbName.Text);
                    break;
                case 2://pdf
                    Document document = new Document();
                    PdfWriter.GetInstance(document, new FileStream("c://123.pdf", FileMode.Create));
                    document.Open();
                    BaseFont bfChinese = BaseFont.CreateFont("C://WINDOWS//Fonts//simsun.ttc,1", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                    iTextSharp.text.Font fontChinese = new iTextSharp.text.Font(bfChinese, 12, iTextSharp.text.Font.NORMAL, new iTextSharp.text.Color(0, 0, 0));

                    //导出文本的内容：
                    document.Add(new Paragraph("你好", fontChinese));
                    //导出图片：
                    //iTextSharp.text.Image jpeg = iTextSharp.text.Image.GetInstance(Path.GetFullPath("1.jpg"));
                    //document.Add(jpeg);

                    //注意一定要关闭，否则PDF中的内容将得不到保存

                    document.Close();
                    break;
            }
        }


    }
}