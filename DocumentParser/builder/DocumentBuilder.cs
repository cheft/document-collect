/***************************************************************************
 * 说明：Office文档构建类
 * 作者：陈海峰
 * 日期：2013-04-26
 *  * 注意事项：
 * 1、依赖于Microsoft Office
 ****************************************************************************/
using DocumentParser.helper;
using log4net;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace DocumentParser.builder
{
    public class DocumentBuilder
    {
        #region 成员变量

        private static readonly ILog log = LogManager.GetLogger(typeof(DocumentBuilder));

        private object objMissing = null;
        private Application word = null;
        private Document doc = null;

        private static string tempDir = ConfigurationManager.AppSettings["QueueDir"] + "\\temp";
        private static string isSheet = ConfigurationManager.AppSettings["IsSheet"];
        private static string isImage = ConfigurationManager.AppSettings["IsImage"];

        public Application getWord()
        {
            return this.word;
        }

        #endregion

        #region 创建空文档
        public DocumentBuilder()
        {
            objMissing = System.Reflection.Missing.Value;
            word = new Application();
            doc = word.Documents.Add(ref objMissing, ref objMissing, ref objMissing, ref objMissing);
            doc.Activate();
        }
        #endregion

        #region 创建模板文档
        public DocumentBuilder(object template)
        {
            objMissing = System.Reflection.Missing.Value;
            object readOnly = false;
            word = new Application();
            object objTempDoc = template;

            if (!File.Exists(template.ToString()))
            {
                doc = word.Documents.Add(ref objMissing, ref objMissing, ref objMissing, ref objMissing);
                // TODO office 2013
                // doc.SaveAs2(template);
                doc.SaveAs(ref objTempDoc, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing);
               }
            else
            {
                doc = word.Documents.Open(
                    ref objTempDoc,   //FileName
                    ref objMissing,   //ConfirmVersions
                    ref readOnly,     //ReadOnly
                    ref objMissing,   //AddToRecentFiles
                    ref objMissing,   //PasswordDocument
                    ref objMissing,   //PasswordTemplate
                    ref objMissing,   //Revert
                    ref objMissing,   //WritePasswordDocument
                    ref objMissing,   //WritePasswordTemplate
                    ref objMissing,   //Format
                    ref objMissing,   //Enconding
                    ref objMissing,   //Visible
                    ref objMissing,   //OpenAndRepair
                    ref objMissing,   //DocumentDirection
                    ref objMissing,   //NoEncodingDialog
                    ref objMissing    //XMLTransform
                    );
            }
            doc.Activate();
            GoToTheEnd();
        }
        #endregion

        #region 缩进 Heading
        public void IndentHeading(int indent)
        {
            bool hasIndex = false;
            // 删除目录 
            if (doc.TablesOfContents.Count > 0 && doc.TablesOfContents[1] != null)
            {
                doc.TablesOfContents[1].Delete();
                hasIndex = true; ;
            }

            foreach (Microsoft.Office.Interop.Word.Paragraph paragraph in doc.Paragraphs)
            {
                Microsoft.Office.Interop.Word.Style style = paragraph.get_Style() as Microsoft.Office.Interop.Word.Style;
                string styleName = style.NameLocal.ToString();

                int i = styleName.IndexOf(",");
                if (i != -1)
                {
                    styleName = styleName.Substring(0, i);
                }
                if (styleName.IndexOf("标题") != -1)
                {
                    int num = 0;
                    try
                    {
                        string level = styleName.Split(' ')[1];
                        num = int.Parse(level);
                    }
                    catch(Exception ex)
                    {
                        log.Error(ex.Message);
                        continue;
                    }
                    // Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading1 = -2;
                    // 其它请看 WdBuiltinStyle 枚举类型,依此得出heading样式值
                    object heading = (object)((-1 - num) - indent);
                    paragraph.set_Style(ref heading);
                }
                else if (hasIndex)
                {
                    string text = paragraph.Range.Text;
                    string regex = "\\s*目\\s*录\\s*";

                    if (Regex.IsMatch(text, regex))
                    {
                        paragraph.Range.Text = "";
                        hasIndex = false;
                    }

                }
            }
        }
        #endregion

        #region 插入对象
        public void InsertObject(string path)
        {
            FileInfo fi = new FileInfo(path);
            string realPath = "";
            if (fi.Length == 0)
            {
                return;
            }
            try
            {
                switch (fi.Extension.ToLower())
                {
                    case ".docx":
                        InsertWord(path);
                        break;
                    case ".doc":
                        InsertWord(path);
                        break;
                    case ".xls":
                        InsertExcel(path);
                        break;
                    case ".xlsx":
                        InsertExcel(path);
                        break;
                    case ".ppt":
                        InsertPpt(path);
                        break;
                    case ".pptx":
                        InsertPpt(path);
                        break;
                        
                    case ".pdf":
                        realPath = path.Substring(0, path.LastIndexOf("."));
                        InsertPdf(realPath);
                        break;
                         
                    case ".png":
                        InsertImage(path);
                        break;
                    case ".jpg":
                        InsertImage(path);
                        break;
                    case ".jepg":
                        InsertImage(path);
                        break;
                    case ".bmp":
                        InsertImage(path);
                        break;
                    case ".gif":
                        InsertImage(path);
                        break;
                    case ".tif":
                        realPath = path.Substring(0, path.LastIndexOf("."));
                        InsertTif(realPath);
                        break;
                    case ".rar":
                        InsertRarPackage(path);
                        break;
                    case ".zip":
                        InsertZipPackage(path);
                        break;
                    case ".txt":
                        InsertContent(ReadTxtFile(path));
                        break;
                    default:
                        AddOLEObject(path);
                        break;
                }
            }
            catch (Exception e)
            {
                AddOLEObject(path);
                Console.WriteLine(e.Message);
            }
            InsertLineBreak();
            
        }

        private void AddOLEObject(string path) 
        {
            object objPath = path;
            word.Selection.InlineShapes.AddOLEObject(ref objMissing, ref objPath, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing);
        }


        private String ReadTxtFile(String path)
        {
            String txt = "";
            using (StreamReader sr = new StreamReader(path, Encoding.Default))
            {
                string s = "";
                while ((s = sr.ReadLine()) != null)
                {
                    txt += s + "\n";
                }
            }
            return txt;
        }

        private void InsertZipPackage(string path)
        {
            string dest = tempDir + "\\" + Guid.NewGuid();
            ZipHelper.UnZip(path, dest);
            string[] files = Directory.GetFiles(dest);
            foreach (string f in files)
            {
                InsertObject(f);
                File.Delete(f);
            }
            Directory.Delete(dest);
        }

        private void InsertRarPackage(string path)
        {
            string dest = tempDir + "\\" + Guid.NewGuid();
            ZipHelper.UnRar(path, dest);
            string[] files = Directory.GetFiles(dest);
            foreach (string f in files)
            {
                InsertObject(f);
                File.Delete(f);
            }
            Directory.Delete(dest);
        }

        #endregion

        #region 插入Word文档
        private void InsertWord(string path)
        {
            if (isImage == "1")
            {
                string tempPdf = tempDir + "\\" + Guid.NewGuid() + ".pdf";
                string outDir = tempDir + "\\" + Guid.NewGuid() + "\\";
                OfficeBuilder ob = new OfficeBuilder();
                ob.Word2PDF(path, tempPdf);
                PdfToImageBuilder ptib = new PdfToImageBuilder();
                ptib.Convert(tempPdf, outDir, "test", PdfToImageBuilder.Definition.One);
                File.Delete(tempPdf);
                string[] images = Directory.GetFiles(outDir);
                foreach (string img in images)
                {
                    InsertImage(img);
                    File.Delete(img);
                }
                Directory.Delete(outDir);
            }
            else
            {
                object objFalse = false;
                object confirmConversion = false;
                object link = false;
                object attachment = false;
                word.Selection.InsertFile(
                    path,
                    ref objMissing,
                    ref confirmConversion,
                    ref link,
                    ref attachment
                );
            }
        }
        #endregion

        #region 插入Excel文档
        public void InsertExcel(string path)
        {
            if (isImage == "1")
            {
                string tempPdf = tempDir + "\\" + Guid.NewGuid() + ".pdf";
                string outDir = tempDir + "\\" + Guid.NewGuid() + "\\";
                OfficeBuilder ob = new OfficeBuilder();
                ob.Excel2PDF(path, tempPdf);
                PdfToImageBuilder ptib = new PdfToImageBuilder();
                ptib.Convert(tempPdf, outDir, "test", PdfToImageBuilder.Definition.One);
                File.Delete(tempPdf);
                string[] images = Directory.GetFiles(outDir);
                foreach (string img in images)
                {
                    InsertImage(img);
                    File.Delete(img);
                }
                Directory.Delete(outDir);
            }
            else
            {
                if (isSheet == "1")
                {
                    string outDir = tempDir + "\\" + Guid.NewGuid();
                    OfficeBuilder ob = new OfficeBuilder();
                    ob.SplitExcel(path, outDir);
                    string[] excels = Directory.GetFiles(outDir);
                    foreach (string e in excels)
                    {
                        AddOLEObject(e);
                        File.Delete(e);
                    }
                    Directory.Delete(outDir);
                }
                else
                {
                    object objPath = path;
                    word.Selection.InlineShapes.AddOLEObject(ref objMissing, ref objPath, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing);
                }
            }
        }
        #endregion

        #region 插入PPT文档
        public void InsertPpt(string path)
        {
            string outDir = tempDir + "\\" + Guid.NewGuid();
            OfficeBuilder ob = new OfficeBuilder();
            ob.PPT2Image(path, outDir);
            string[] images = Directory.GetFiles(outDir);
            foreach (string img in images)
            {
                InsertImage(img);
                File.Delete(img);
            }
            Directory.Delete(outDir);
        }
        #endregion

        #region 插入PDF文档
        private void InsertPdf(string imageDir)
        {
            string[] images = Directory.GetFiles(imageDir);
            foreach (string img in images)
            {
                InsertImage(img);
                // File.Delete(img);
            }
            Directory.Delete(imageDir, true);        
        }
        #endregion

        #region 插入TIF文档
        private void InsertTif(string imageDir)
        {
            InsertPdf(imageDir);
        }
        #endregion

        #region 插入图片
        private void InsertImage(string path)
        {
            word.Selection.InlineShapes.AddPicture(path, ref objMissing, ref objMissing, ref objMissing);
        }
        #endregion

        #region 保存文件
        public void SaveAs(string outDoc)
        {
            object objOutDoc = outDoc;
            try
            {
                doc.SaveAs(
                  ref objOutDoc,      //FileName
                  ref objMissing,     //FileFormat
                  ref objMissing,     //LockComments
                  ref objMissing,     //PassWord    
                  ref objMissing,     //AddToRecentFiles
                  ref objMissing,     //WritePassword
                  ref objMissing,     //ReadOnlyRecommended
                  ref objMissing,     //EmbedTrueTypeFonts
                  ref objMissing,     //SaveNativePictureFormat
                  ref objMissing,     //SaveFormsData
                  ref objMissing,     //SaveAsAOCELetter,
                  ref objMissing,     //Encoding
                  ref objMissing,     //InsertLineBreaks
                  ref objMissing,     //AllowSubstitutions
                  ref objMissing,     //LineEnding
                  ref objMissing      //AddBiDiMarks
                  );
            }
            finally
            {
                word.Quit(
                  ref objMissing,     //SaveChanges
                  ref objMissing,     //OriginalFormat
                  ref objMissing      //RoutDocument
                  );
                word = null;
                WordHelper.KillWordProcess();
            }
        }
        #endregion

        #region 替换封面内容
        public void InsertFrontCover(string s)
        {
            string[] arr = s.Split('@');
            foreach (string a in arr)
            {
                string[] sp = a.Split('%');
                foreach (Microsoft.Office.Interop.Word.Bookmark bm in doc.Bookmarks)
               {
                    if (bm.Name == sp[1])
                    {
                        bm.Select();
                        bm.Range.Text = sp[2];
                    }
                }
            }
        }
        #endregion

        #region 给word文档添加页眉页脚 无用
        /// <summary>
        /// 给word文档添加页眉
        /// </summary>
        /// <param name="filePath">文件名</param>
        /// <returns></returns>
        public bool AddPageHeaderFooter(string filePath)
        {
            try
            {
                Object oMissing = System.Reflection.Missing.Value;
                word.Visible = true;
                object filename = filePath;
                
                ////添加页眉方法一：
                //WordApp.ActiveWindow.View.Type = WdViewType.wdOutlineView;
                //WordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekPrimaryHeader;
                //WordApp.ActiveWindow.ActivePane.Selection.InsertAfter( "**公司" );//页眉内容

                ////添加页眉方法二：
                if (word.ActiveWindow.ActivePane.View.Type == WdViewType.wdNormalView ||
                    word.ActiveWindow.ActivePane.View.Type == WdViewType.wdOutlineView)
                {
                    word.ActiveWindow.ActivePane.View.Type = WdViewType.wdPrintView;
                }
                word.ActiveWindow.View.SeekView = WdSeekView.wdSeekCurrentPageHeader;
                word.Selection.HeaderFooter.LinkToPrevious = false;
                word.Selection.HeaderFooter.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                word.Selection.HeaderFooter.Range.Text = "页眉内容";

                word.ActiveWindow.View.SeekView = WdSeekView.wdSeekCurrentPageFooter;
                word.Selection.HeaderFooter.LinkToPrevious = false;
                word.Selection.HeaderFooter.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                word.ActiveWindow.ActivePane.Selection.InsertAfter("页脚内容");

                //跳出页眉页脚设置
                word.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument;

                return true;
            }
            catch (Exception e)
            {
                log.Error(e.Message);
                return false;
            }
        }
        #endregion 给word文档添加页眉页脚

        #region 建立目录
        public void BuildContentsIndex()
        {
            Object oTrue = true;
            Object oFalse = false;
            word.Visible = true;

            //---------------------------------------------------------------------------------------------------------------------
            word.Selection.Paragraphs.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            word.Selection.Paragraphs.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            word.Selection.Paragraphs.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;

            Range myRange = null;

            Object oUpperHeadingLevel = "1";
            Object oLowerHeadingLevel = "3";
            
          
            foreach (Microsoft.Office.Interop.Word.Bookmark bm in doc.Bookmarks)
            {
                if (bm.Name == "IndexBookmark")
                {
                    bm.Select();
                    myRange = bm.Range;
                }
               
            }

            doc.TablesOfContents.Add(myRange, ref oTrue, ref oUpperHeadingLevel,
                ref oLowerHeadingLevel, ref objMissing, ref objMissing, ref oTrue,
                ref oTrue, ref objMissing, ref oTrue, ref oTrue, ref oTrue);
        }
        #endregion

        #region 插入书签
        public void InsertBookmark(string key)
        {
            doc.Bookmarks.Add(key, ref objMissing);
        }
        #endregion

        #region 关闭word进程
        /// <summary>
        /// 关闭word进程
        /// </summary>
        public void KillWordProcess()
        {
            System.Diagnostics.Process[] myProcess;
            myProcess = System.Diagnostics.Process.GetProcesses();
            foreach (System.Diagnostics.Process process in myProcess)
            {
                if (process.Id != 0)
                {
                    string myS = "WINWORD.EXE" + process.ProcessName + "  ID:" + process.Id.ToString();
                    try
                    {
                        if (process.Modules != null)
                            if (process.Modules.Count > 0)
                            {
                                System.Diagnostics.ProcessModule pm = process.Modules[0];
                                myS += "\n Modules[0].FileName:" + pm.FileName;
                                myS += "\n Modules[0].ModuleName:" + pm.ModuleName;
                                myS += "\n Modules[0].FileVersionInfo:\n" + pm.FileVersionInfo.ToString();
                                if (pm.ModuleName.ToLower() == "winword.exe")
                                    process.Kill();
                            }
                    }
                    catch(Exception ex)
                    {
                        log.Error(ex.Message);
                    }
                    finally
                    {
                    }
                }
            }
        }
        #endregion 关闭word进程

        #region 判断系统是否装word

        /// <summary>
        /// 判断系统是否装word
        /// </summary>
        /// <returns></returns>
        public bool IsWordInstalled()
        {
            RegistryKey machineKey = Registry.LocalMachine;
            if (IsWordInstalledByVersion("12.0", machineKey))
            {
                return true;
            }
            if (IsWordInstalledByVersion("11.0", machineKey))
            {
                return true;
            }
            return false;
        }


        /// <summary>
        /// 判断系统是否装某版本的word
        /// </summary>
        /// <param name="strVersion">版本号</param>
        /// <param name="machineKey"></param>
        /// <returns></returns>
        private bool IsWordInstalledByVersion(string strVersion, RegistryKey machineKey)
        {
            try
            {
                RegistryKey installKey =
                    machineKey.OpenSubKey("Software").OpenSubKey("Microsoft").OpenSubKey(
                    "Office").OpenSubKey(strVersion).OpenSubKey("Word").OpenSubKey("InstallRoot");
                if (installKey == null)
                {
                    return false;
                }
                return true;
            }
            catch (Exception e)
            {
                log.Error(e.Message);
                return false;
            }
        }
        #endregion 判断系统是否装word

        #region 文件操作

        // Open a file (the file must exists) and activate it
        public void Open(string strFileName)
        {
            object fileName = strFileName;
            object readOnly = false;
            object isVisible = true;

            doc = word.Documents.Open(ref fileName, ref objMissing, ref readOnly,
                ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing,
                ref objMissing, ref objMissing, ref isVisible, ref objMissing, ref objMissing, ref objMissing, ref objMissing);

            doc.Activate();
        }

        // Open a new document
        public void Open()
        {
            doc = word.Documents.Add(ref objMissing, ref objMissing, ref objMissing, ref objMissing);

            doc.Activate();
        }

        public void Quit()
        {
            word.Application.Quit(ref objMissing, ref objMissing, ref objMissing);
        }

        public void Save()
        {
            doc.Save();
        }


        // Save the document in HTML format
        public void SaveAsHtml(string strFileName)
        {
            object fileName = strFileName;
            object Format = (int)Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatHTML;
            doc.SaveAs(ref fileName, ref Format, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing,
                ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing);
        }

        public void SaveAsPDF(string strFileName)
        {
            // TODO office 2013
            /*
            object fileName = strFileName;
            object Format = (int)Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF;
            doc.SaveAs(ref fileName, ref Format, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing,
                ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing);
            */

        }

        #endregion

        #region 移动光标位置

        // Go to a predefined bookmark, if the bookmark doesn't exists the application will raise an error
        public void GotoBookMark(string strBookMarkName)
        {
            // VB :  Selection.GoTo What:=wdGoToBookmark, Name:="nome"
            object Bookmark = (int)Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark;
            object NameBookMark = strBookMarkName;
            word.Selection.GoTo(ref Bookmark, ref objMissing, ref objMissing, ref NameBookMark);
        }

        public void GoToTheEnd()
        {
            // VB :  Selection.EndKey Unit:=wdStory
            object unit;
            unit = Microsoft.Office.Interop.Word.WdUnits.wdStory;
            word.Selection.EndKey(ref unit, ref objMissing);
        }

        public void GoToLineEnd()
        {
            object unit = Microsoft.Office.Interop.Word.WdUnits.wdLine;
            object ext = Microsoft.Office.Interop.Word.WdMovementType.wdExtend;
            word.Selection.EndKey(ref unit, ref ext);
        }

        public void GoToTheBeginning()
        {
            // VB : Selection.HomeKey Unit:=wdStory
            object unit;
            unit = Microsoft.Office.Interop.Word.WdUnits.wdStory;
            word.Selection.HomeKey(ref unit, ref objMissing);
        }

        public void GoToRightCell()
        {
            // Selection.MoveRight Unit:=wdCell
            object direction;
            direction = Microsoft.Office.Interop.Word.WdUnits.wdCell;
            word.Selection.MoveRight(ref direction, ref objMissing, ref objMissing);
        }

        public void GoToLeftCell()
        {
            // Selection.MoveRight Unit:=wdCell
            object direction;
            direction = Microsoft.Office.Interop.Word.WdUnits.wdCell;
            word.Selection.MoveLeft(ref direction, ref objMissing, ref objMissing);
        }

        public void GoToDownCell()
        {
            // Selection.MoveRight Unit:=wdCell
            object direction;
            direction = Microsoft.Office.Interop.Word.WdUnits.wdLine;
            word.Selection.MoveDown(ref direction, ref objMissing, ref objMissing);
        }

        public void GoToUpCell()
        {
            // Selection.MoveRight Unit:=wdCell
            object direction;
            direction = Microsoft.Office.Interop.Word.WdUnits.wdLine;
            word.Selection.MoveUp(ref direction, ref objMissing, ref objMissing);
        }

        #endregion

        #region 插入操作

        public void InsertTitle(string text)
        {
            SetAlignment("center");
            SetFont("Bold");
            SetFontSize(30);
            // TODO office 2013
            // object styleType = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleBookTitle;
            // InsertText(text, styleType);

            // TODO office 2003
            word.Selection.TypeText(text);
            InsertLineBreak();

        }

        public void InsertContent(string text)
        {
            object styleType = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleBodyText;
            InsertText(text, styleType);
        }

        private void InsertText(string text, object styleType)
        {
            word.ActiveWindow.Selection.Range.set_Style(ref styleType);
            word.Selection.TypeText(text);
            InsertLineBreak();
        }

        public void InsertHeading(int level, string text)
        {
            // Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading1 = -2;
            // 其它请看 WdBuiltinStyle 枚举类型,依此得出heading样式值
            object headingStyle = (object)(-1 - level);
            word.ActiveWindow.Selection.Range.set_Style(ref headingStyle);
            word.Selection.TypeText(text);
            word.Selection.TypeParagraph();
        }

        public void InsertHeading1(string text)
        {
            object headingType = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading1;
            InsertText(text, headingType);
        }

        public void InsertHeading2(string text)
        {
            object headingType = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading2;
            InsertText(text, headingType);
        }

        public void InsertHeading3(string text)
        {
            object headingType = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading3;
            InsertText(text, headingType);
        }

        public void InsertHeading4(string text)
        {
            object headingType = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading4;
            InsertText(text, headingType);
        }

        public void InsertHeading5(string text)
        {
            object headingType = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading5;
            InsertText(text, headingType);
        }

        public void InsertHeading6(string text)
        {
            object headingType = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading6;
            InsertText(text, headingType);
        }

        public void InsertHeading7(string text)
        {
            object headingType = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading7;
            InsertText(text, headingType);
        }

        public void InsertHeading8(string text)
        {
            object headingType = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading8;
            InsertText(text, headingType);
        }

        public void InsertHeading9(string text)
        {
            object headingType = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleHeading9;
            InsertText(text, headingType);
        }

        public void InsertLineBreak()
        {
            word.Selection.TypeParagraph();
        }

        /// <summary>
        /// 插入多个空行
        /// </summary>
        /// <param name="nline"></param>
        public void InsertLineBreak(int nline)
        {
            for (int i = 0; i < nline; i++)
                word.Selection.TypeParagraph();
        }

        public void InsertPageBreak()
        {
            // VB : Selection.InsertBreak Type:=wdPageBreak
            object pBreak = (int)Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
            word.Selection.InsertBreak(ref pBreak);
        }

        // 插入页码
        public void InsertPageNumber()
        {
            object wdFieldPage = Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage;
            object preserveFormatting = true;
            word.Selection.Fields.Add(word.Selection.Range, ref wdFieldPage, ref objMissing, ref preserveFormatting);
        }

        // 插入页码
        public void InsertPageNumber(string strAlign)
        {
            object wdFieldPage = Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage;
            object preserveFormatting = true;
            word.Selection.Fields.Add(word.Selection.Range, ref wdFieldPage, ref objMissing, ref preserveFormatting);
            SetAlignment(strAlign);
        }

        public void InsertImage(string strPicPath, float picWidth, float picHeight)
        {
            string FileName = strPicPath;
            object LinkToFile = false;
            object SaveWithDocument = true;
            object Anchor = word.Selection.Range;
            word.ActiveDocument.InlineShapes.AddPicture(FileName, ref LinkToFile, ref SaveWithDocument, ref Anchor).Select();
            word.Selection.InlineShapes[1].Width = picWidth; // 图片宽度 
            word.Selection.InlineShapes[1].Height = picHeight; // 图片高度

            // 将图片设置为四面环绕型 
            Microsoft.Office.Interop.Word.Shape s = word.Selection.InlineShapes[1].ConvertToShape();
            s.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapSquare;
        }

        public void InsertLine(float left, float top, float width, float weight, int r, int g, int b)
        {
            //SetFontColor("red");
            //SetAlignment("Center");
            object Anchor = word.Selection.Range;
            //int pLeft = 0, pTop = 0, pWidth = 0, pHeight = 0;
            //word.ActiveWindow.GetPoint(out pLeft, out pTop, out pWidth, out pHeight,objMissing);
            //MessageBox.Show(pLeft + "," + pTop + "," + pWidth + "," + pHeight);
            object rep = false;
            //left += word.ActiveDocument.PageSetup.LeftMargin;
            left = word.CentimetersToPoints(left);
            top = word.CentimetersToPoints(top);
            width = word.CentimetersToPoints(width);
            Microsoft.Office.Interop.Word.Shape s = word.ActiveDocument.Shapes.AddLine(0, top, width, top, ref Anchor);
            s.Line.ForeColor.RGB = RGB(r, g, b);
            s.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            s.Line.Style = Microsoft.Office.Core.MsoLineStyle.msoLineSingle;
            s.Line.Weight = weight;
        }

        #endregion

        #region 设置样式

        /// <summary>
        /// Change the paragraph alignement
        /// </summary>
        /// <param name="strType"></param>
        public void SetAlignment(string strType)
        {
            switch (strType.ToLower())
            {
                case "center":
                    word.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    break;
                case "left":
                    word.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    break;
                case "right":
                    word.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;
                    break;
                case "justify":
                    word.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                    break;
            }

        }


        // if you use thif function to change the font you should call it again with 
        // no parameter in order to set the font without a particular format
        public void SetFont(string strType)
        {
            switch (strType)
            {
                case "Bold":
                    word.Selection.Font.Bold = 1;
                    break;
                case "Italic":
                    word.Selection.Font.Italic = 1;
                    break;
                case "Underlined":
                    word.Selection.Font.Subscript = 0;
                    break;
            }
        }

        // disable all the style 
        public void SetFont()
        {
            word.Selection.Font.Bold = 0;
            word.Selection.Font.Italic = 0;
            word.Selection.Font.Subscript = 0;

        }

        public void SetFontName(string strType)
        {
            word.Selection.Font.Name = strType;
        }

        public void SetFontSize(float nSize)
        {
            SetFontSize(nSize, 100);
        }

        public void SetFontSize(float nSize, int scaling)
        {
            if (nSize > 0f)
                word.Selection.Font.Size = nSize;
            if (scaling > 0)
                word.Selection.Font.Scaling = scaling;
        }

        public void SetFontColor(string strFontColor)
        {
            switch (strFontColor.ToLower())
            {
                case "blue":
                    word.Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlue;
                    break;
                case "gold":
                    word.Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorGold;
                    break;
                case "gray":
                    word.Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorGray875;
                    break;
                case "green":
                    word.Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorGreen;
                    break;
                case "lightblue":
                    word.Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorLightBlue;
                    break;
                case "orange":
                    word.Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorOrange;
                    break;
                case "pink":
                    word.Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorPink;
                    break;
                case "red":
                    word.Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed;
                    break;
                case "yellow":
                    word.Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorYellow;
                    break;
            }
        }

        public void SetPageNumberAlign(string strType, bool bHeader)
        {
            object alignment;
            object bFirstPage = false;
            object bF = true;
            //if (bHeader == true)
            //WordApplic.Selection.HeaderFooter.PageNumbers.ShowFirstPageNumber = bF;
            switch (strType)
            {
                case "Center":
                    alignment = Microsoft.Office.Interop.Word.WdPageNumberAlignment.wdAlignPageNumberCenter;
                    //WordApplic.Selection.HeaderFooter.PageNumbers.Add(ref alignment,ref bFirstPage);
                    //Microsoft.Office.Interop.Word.Selection objSelection = WordApplic.pSelection;
                    word.Selection.HeaderFooter.PageNumbers[1].Alignment = Microsoft.Office.Interop.Word.WdPageNumberAlignment.wdAlignPageNumberCenter;
                    break;
                case "Right":
                    alignment = Microsoft.Office.Interop.Word.WdPageNumberAlignment.wdAlignPageNumberRight;
                    word.Selection.HeaderFooter.PageNumbers[1].Alignment = Microsoft.Office.Interop.Word.WdPageNumberAlignment.wdAlignPageNumberRight;
                    break;
                case "Left":
                    alignment = Microsoft.Office.Interop.Word.WdPageNumberAlignment.wdAlignPageNumberLeft;
                    word.Selection.HeaderFooter.PageNumbers.Add(ref alignment, ref bFirstPage);
                    break;
            }
        }

        /// <summary>
        /// 设置页面为标准A4公文样式
        /// </summary>
        private void SetA4PageSetup()
        {
            word.ActiveDocument.PageSetup.TopMargin = word.CentimetersToPoints(3.7f);
            //word.ActiveDocument.PageSetup.BottomMargin = word.CentimetersToPoints(1f);
            word.ActiveDocument.PageSetup.LeftMargin = word.CentimetersToPoints(2.8f);
            word.ActiveDocument.PageSetup.RightMargin = word.CentimetersToPoints(2.6f);
            //word.ActiveDocument.PageSetup.HeaderDistance = word.CentimetersToPoints(2.5f);
            //word.ActiveDocument.PageSetup.FooterDistance = word.CentimetersToPoints(1f);
            word.ActiveDocument.PageSetup.PageWidth = word.CentimetersToPoints(21f);
            word.ActiveDocument.PageSetup.PageHeight = word.CentimetersToPoints(29.7f);
        }

        #endregion

        #region 替换

        ///<summary>
        /// 在word 中查找一个字符串直接替换所需要的文本
        /// </summary>
        /// <param name="strOldText">原文本</param>
        /// <param name="strNewText">新文本</param>
        /// <param name="replaceType">0 不替换找到的任何项, 1 替换找到的第一项, 2 *wdReplaceAll - 替换找到的所有项</param>// 
        /// <returns></returns>
        public bool Replace(string strOldText, string strNewText, int replaceType)
        {
            if (doc == null)
                doc = word.ActiveDocument;
            doc.Content.Find.Text = strOldText;
            object FindText, ReplaceWith, Replace; 
            FindText = strOldText;//要查找的文本
            ReplaceWith = strNewText;//替换文本
            Replace = (object)replaceType;
            doc.Content.Find.ClearFormatting();//移除Find的搜索文本和段落格式设置
            if (doc.Content.Find.Execute(
                ref FindText, ref objMissing,
                ref objMissing, ref objMissing,
                ref objMissing, ref objMissing,
                ref objMissing, ref objMissing, ref objMissing,
                ref ReplaceWith, ref Replace,
                ref objMissing, ref objMissing,
                ref objMissing, ref objMissing))
            {
                return true;
            }
            return false;
        }

        #endregion

        #region RGB颜色互转函数
        /// <summary>
        /// rgb转换函数
        /// </summary>
        /// <param name="r"></param>
        /// <param name="g"></param>
        /// <param name="b"></param>
        /// <returns></returns>
        int RGB(int r, int g, int b)
        {
            return ((b << 16) | (ushort)(((ushort)g << 8) | r));
        }

        Color RGBToColor(int color)
        {
            int r = 0xFF & color;
            int g = 0xFF00 & color;
            g >>= 8;
            int b = 0xFF0000 & color;
            b >>= 16;
            return Color.FromArgb(r, g, b);
        }
        #endregion

    }
}
