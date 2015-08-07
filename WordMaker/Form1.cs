using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace WordMaker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private string ShowXml(string str)
        {
            try
            {
                MemoryStream mstream = new MemoryStream();
                XmlTextWriter writer = new XmlTextWriter(mstream, null);
                XmlDocument xmldoc = new XmlDocument();
                writer.Formatting = Formatting.Indented;

                xmldoc.LoadXml(str);
                xmldoc.WriteTo(writer);
                writer.Flush();
                writer.Close();

                Encoding encoding = Encoding.GetEncoding("utf-8");
                string strReturn = encoding.GetString(mstream.ToArray());
                mstream.Close();
                return strReturn;
            }
            catch
            {
                return string.Empty;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            WordMaker wm = new WordMaker();
            List<IOTOP.Model.Sys_Sdk> list = new IOTOP.BLL.Sys_SdkManager().GetListByUser("SDK002");
            if (list.Count == 0)
                return;
            int titie1, title2;
            titie1 = title2 = 0;
            try
            {
                wm.CreateAWord();
                wm.SetPageHeader("物联网智能开放共性平台说明文档");
                wm.SetTableofContents();
                foreach (IOTOP.Model.Sys_Sdk model in list)
                {
                    if (model.PREFUNCID == "")
                    {
                        titie1++;
                        title2 = 0;
                        wm.InsertText("1 " + model.SdkName, float.Parse("16"), "宋体", Microsoft.Office.Interop.Word.WdColor.wdColorBlack, 0, Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft);
                        wm.NewLine();
                    }
                    else
                    {
                        title2++;
                        string node = titie1.ToString() + "." + title2.ToString();
                        //接口名 宋体 12
                        wm.InsertText("    " + node + " " + model.SdkName, float.Parse("12"), "宋体", Microsoft.Office.Interop.Word.WdColor.wdColorBlack, 0, Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft);
                        wm.NewLine();

                        //功能描述 宋体 10.5
                        wm.InsertText("        " + node + ".1 功能描述", float.Parse("10.5"), "宋体", Microsoft.Office.Interop.Word.WdColor.wdColorBlack, 0, Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft);
                        wm.NewLine();
                        //功能描述内容 宋体 10
                        wm.InsertText("            " + model.Sdkinfo, float.Parse("10"), "宋体", Microsoft.Office.Interop.Word.WdColor.wdColorBlack, 0, Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft);
                        wm.NewLine();

                        //请求参数列表 宋体 10.5
                        wm.InsertText("        " + node + ".2 请求参数列表", float.Parse("10.5"), "宋体", Microsoft.Office.Interop.Word.WdColor.wdColorBlack, 0, Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft);

                        //插入参数表格
                        wm.InsertTable(model.Sdkargs.Split(';'));

                        wm.InsertText("        " + node + ".3 返回XML释义", float.Parse("10.5"), "宋体", Microsoft.Office.Interop.Word.WdColor.wdColorBlack, 0, Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft);
                        wm.NewLine();
                        //返回XML释义
                        wm.InsertText(ShowXml(model.Returns), float.Parse("10"), "宋体", Microsoft.Office.Interop.Word.WdColor.wdColorBlack, 0, Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft);

                        wm.NewLine();
                    }
                }


                //wm.InsertTable();
                wm.SaveWord("D:/5.doc");
            }
            finally
            {
                wm.Dispose();
            }
        }
    }
}
