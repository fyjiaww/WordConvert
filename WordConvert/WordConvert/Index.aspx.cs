using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;

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
                return "0";//没有文件
            }

            Microsoft.Office.Interop.Word.ApplicationClass word = new Microsoft.Office.Interop.Word.ApplicationClass();
            Type wordType = word.GetType();
            Microsoft.Office.Interop.Word.Documents docs = word.Documents;

            // 打开文件  
            Type docsType = docs.GetType();

            object fileName = strFile;

            Microsoft.Office.Interop.Word.Document doc = (Microsoft.Office.Interop.Word.Document)docsType.InvokeMember("Open",
            System.Reflection.BindingFlags.InvokeMethod, null, docs, new Object[] { fileName, true, true });

            // 转换格式，另存为html  
            Type docType = doc.GetType();
            //给文件重新起名
            string filename = System.DateTime.Now.Year.ToString() + System.DateTime.Now.Month.ToString() + System.DateTime.Now.Day.ToString() +
            System.DateTime.Now.Hour.ToString() + System.DateTime.Now.Minute.ToString() + System.DateTime.Now.Second.ToString();

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

            File.Delete(strFile);//删除临时保存的文件

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
                html = System.Text.RegularExpressions.Regex.Replace(html, @"<meta[^>]*>", "<meta http-equiv=Content-Type content='text/html; charset=gb2312'>", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
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

        //导出按钮
        protected void Button3_Click(object sender, EventArgs e)
        {
            switch (this.DropDownList1.SelectedIndex)
            {
                case 0://word

                    break;
                case 1://excel
                    NPOI.HSSF.UserModel.HSSFWorkbook book = new NPOI.HSSF.UserModel.HSSFWorkbook();
                    NPOI.SS.UserModel.ISheet sheet = book.CreateSheet("test_01");

                    // 第一列
                    NPOI.SS.UserModel.IRow row = sheet.CreateRow(0);
                    row.CreateCell(0).SetCellValue("第一列第一行");

                    // 第二列
                    NPOI.SS.UserModel.IRow row2 = sheet.CreateRow(1);
                    row2.CreateCell(0).SetCellValue("第二列第一行");

                    // 写入到客户端  
                    System.IO.MemoryStream ms = new System.IO.MemoryStream();
                    book.Write(ms);
                    Response.AddHeader("Content-Disposition", string.Format("attachment; filename={0}.xls", DateTime.Now.ToString("yyyyMMddHHmmssfff")));
                    Response.BinaryWrite(ms.ToArray());
                    book = null;
                    ms.Close();
                    ms.Dispose();
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

        public class ListTemp
        {
            public int Id { get; set; }
            public string Name { get; set; }

        }
    }
}