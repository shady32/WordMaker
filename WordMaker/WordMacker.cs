using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;

namespace WordMaker
{
    public class WordMaker
    {
        private Microsoft.Office.Interop.Word.Application _wordApplication;
        private Microsoft.Office.Interop.Word.Document _wordDocument;

        public WordMaker()
        {

        }

        public void CreateAWord()
        {
            //实例化word应用对象    
            this._wordApplication = new Microsoft.Office.Interop.Word.ApplicationClass();
            Object myNothing = System.Reflection.Missing.Value;

            this._wordDocument = this._wordApplication.Documents.Add(ref myNothing, ref myNothing, ref myNothing, ref myNothing);
        }

        public void Dispose()
        {
            this._wordApplication.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(this._wordApplication);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(this._wordDocument);
            this._wordApplication = null;
            this._wordDocument = null;
            GC.Collect();
            GC.Collect();
        }

        public void SetTableofContents()
        {
            this._wordDocument.TablesOfContents.Add(this._wordApplication.Selection.Range);
        }

        public void SetPageHeader(string pPageHeader)
        {
            //添加页眉    
            this._wordApplication.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdOutlineView;
            this._wordApplication.ActiveWindow.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekPrimaryHeader;
            this._wordApplication.ActiveWindow.ActivePane.Selection.InsertAfter(pPageHeader);
            //设置中间对齐    
            this._wordApplication.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            //跳出页眉设置    
            this._wordApplication.ActiveWindow.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument;
        }

        /// <SUMMARY></SUMMARY>    
        /// 插入文字    
        ///     
        /// <PARAM name="pText" />文本信息    
        /// <PARAM name="pFontSize" />字体打小    
        /// <PARAM name="pFontColor" />字体颜色    
        /// <PARAM name="pFontBold" />字体粗体    
        /// <PARAM name="ptextAlignment" />方向    
        public void InsertText(string pText, float pFontSize, string familyName, Microsoft.Office.Interop.Word.WdColor pFontColor, int pFontBold, Microsoft.Office.Interop.Word.WdParagraphAlignment ptextAlignment)
        {
            //设置字体样式以及方向    
            this._wordApplication.Application.Selection.Font.Size = pFontSize;
            this._wordApplication.Application.Selection.Font.Bold = pFontBold;
            this._wordApplication.Application.Selection.Font.Color = pFontColor;
            this._wordApplication.Application.Selection.ParagraphFormat.Alignment = ptextAlignment;
            this._wordApplication.Application.Selection.Font.Name = familyName;
            this._wordApplication.Application.Selection.TypeText(pText);
        }

        public void NewLine()
        {
            //换行    
            this._wordApplication.Application.Selection.TypeParagraph();
        }
        public void InsertPicture(string pPictureFileName)
        {
            object myNothing = System.Reflection.Missing.Value;
            //图片居中显示    
            this._wordApplication.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            this._wordApplication.Application.Selection.InlineShapes.AddPicture(pPictureFileName, ref myNothing, ref myNothing, ref myNothing);
        }

        public void InsertTable(string[] args)
        {
            this._wordApplication.Application.Selection.Tables.Add(this._wordApplication.Selection.Range, args.Length + 1, 6);

            #region tables
            int i = this._wordDocument.Application.Selection.Tables.Count;
            this._wordApplication.Application.Selection.Tables[i].Cell(1, 1).Range.InsertAfter("参数");
            this._wordApplication.Application.Selection.Tables[i].Cell(1, 2).Range.InsertAfter("参数名称");
            this._wordApplication.Application.Selection.Tables[i].Cell(1, 3).Range.InsertAfter("类型");
            this._wordApplication.Application.Selection.Tables[i].Cell(1, 4).Range.InsertAfter("参数说明");
            this._wordApplication.Application.Selection.Tables[i].Cell(1, 5).Range.InsertAfter("是否可为空");
            this._wordApplication.Application.Selection.Tables[i].Cell(1, 6).Range.InsertAfter("样例");

            this._wordApplication.Application.Selection.Tables[i].Cell(1, 1).Range.Shading.BackgroundPatternColor
                = Microsoft.Office.Interop.Word.WdColor.wdColorSkyBlue;
            this._wordApplication.Application.Selection.Tables[i].Cell(1, 2).Range.Shading.BackgroundPatternColor
                = Microsoft.Office.Interop.Word.WdColor.wdColorSkyBlue;
            this._wordApplication.Application.Selection.Tables[i].Cell(1, 3).Range.Shading.BackgroundPatternColor
                = Microsoft.Office.Interop.Word.WdColor.wdColorSkyBlue;
            this._wordApplication.Application.Selection.Tables[i].Cell(1, 4).Range.Shading.BackgroundPatternColor
                = Microsoft.Office.Interop.Word.WdColor.wdColorSkyBlue;
            this._wordApplication.Application.Selection.Tables[i].Cell(1, 5).Range.Shading.BackgroundPatternColor
                = Microsoft.Office.Interop.Word.WdColor.wdColorSkyBlue;
            this._wordApplication.Application.Selection.Tables[i].Cell(1, 6).Range.Shading.BackgroundPatternColor
                = Microsoft.Office.Interop.Word.WdColor.wdColorSkyBlue;

            this._wordApplication.Selection.Tables[i].Select();
            this._wordApplication.Selection.Tables[i].Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft].LineStyle
                = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            this._wordApplication.Selection.Tables[i].Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft].LineWidth
                = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth025pt;
            this._wordApplication.Selection.Tables[i].Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft].ColorIndex
                = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkBlue;

            this._wordApplication.Selection.Tables[i].Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight].LineStyle
                = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            this._wordApplication.Selection.Tables[i].Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight].LineWidth
                = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth025pt;
            this._wordApplication.Selection.Tables[i].Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight].ColorIndex
                = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkBlue;

            this._wordApplication.Selection.Tables[i].Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop].LineStyle
                = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            this._wordApplication.Selection.Tables[i].Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop].LineWidth
                = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth025pt;
            this._wordApplication.Selection.Tables[i].Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop].ColorIndex
                = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkBlue;

            this._wordApplication.Selection.Tables[i].Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom].LineStyle
                = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            this._wordApplication.Selection.Tables[i].Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom].LineWidth
                = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth025pt;
            this._wordApplication.Selection.Tables[i].Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom].ColorIndex
                = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkBlue;

            this._wordApplication.Selection.Tables[i].Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal].LineStyle
                = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            this._wordApplication.Selection.Tables[i].Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal].LineWidth
                = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth025pt;
            this._wordApplication.Selection.Tables[i].Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal].ColorIndex
                = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkBlue;

            this._wordApplication.Selection.Tables[i].Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical].LineStyle
                = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
            this._wordApplication.Selection.Tables[i].Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical].LineWidth
                = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth025pt;
            this._wordApplication.Selection.Tables[i].Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical].ColorIndex
                = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkBlue;

            #endregion

            for (int cc = 0; cc < args.Length; cc++)
            {
                string[] arg = args[cc].Split(',');
                int colums = 1;
                while (colums < 7)
                {
                    try
                    {
                        this._wordApplication.Application.Selection.Tables[i].Cell(cc + 2, colums).Range.InsertAfter(arg[colums - 1]);
                    }
                    catch
                    {
                        this._wordApplication.Application.Selection.Tables[i].Cell(cc + 2, colums).Range.InsertAfter("");
                    }
                    finally
                    {
                        colums++;
                    }
                }
                //this._wordApplication.Application.Selection.Tables[i].Cell(cc + 2, 1).Range.InsertAfter(arg[0]);
                //this._wordApplication.Application.Selection.Tables[i].Cell(cc + 2, 2).Range.InsertAfter(arg[1]);
                //this._wordApplication.Application.Selection.Tables[i].Cell(cc + 2, 3).Range.InsertAfter(arg[2]);
                //this._wordApplication.Application.Selection.Tables[i].Cell(cc + 2, 4).Range.InsertAfter(arg[3]);
                //this._wordApplication.Application.Selection.Tables[i].Cell(cc + 2, 5).Range.InsertAfter(arg[4]);
                //this._wordApplication.Application.Selection.Tables[i].Cell(cc + 2, 6).Range.InsertAfter(arg[5]);
                this._wordApplication.Selection.MoveDown();
            }

            for (int c1 = 0; c1 < args.Length + 1; c1++)
            {
                for (int c2 = 0; c2 < 6; c2++)
                {
                    this._wordApplication.Application.Selection.Tables[i].Cell(c1 + 1, c2 + 1).Range.Font.Size = 9;
                }
            }
        }

        public void SaveWord(string pFileName)
        {
            object myNothing = System.Reflection.Missing.Value;
            object myFileName = pFileName;
            object myWordFormatDocument = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocument;
            object myLockd = false;
            object myPassword = "";
            object myAddto = true;
            try
            {
                this._wordDocument.SaveAs(ref myFileName, ref myWordFormatDocument, ref myLockd, ref myPassword, ref myAddto, ref myPassword,
                    ref myLockd, ref myLockd, ref myLockd, ref myLockd, ref myNothing, ref myNothing, ref myNothing,
                    ref myNothing, ref myNothing, ref myNothing);
            }
            catch (Exception ex)
            {
                throw new Exception("导出word文档失败!");
            }
        }
    }
}