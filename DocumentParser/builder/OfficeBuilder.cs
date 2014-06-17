using System.Reflection;
using System;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using log4net;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace DocumentParser.builder
{
    public class OfficeBuilder
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(OfficeBuilder));

        /// <summary>
        /// EXCEL文檔轉成PDF文檔
        ///  參考 http://msdn.microsoft.com/en-us/library/bb256835%28v=office.12%29.aspx
        ///  Open XML SDK 2.0 for Microsoft Office http://www.microsoft.com/en-us/download/details.aspx?id=5124
        ///  
        /// </summary>
        /// <param name="infile"></param>
        /// <param name="outfile"></param>
       public  void Excel2PDF(string infile, string outfile)
        {
            object objOpt = Missing.Value;
            Excel.Application excelApp = null;
            try
            {
                excelApp = new Excel.Application();
                excelApp.Workbooks.Open(infile, objOpt, objOpt, objOpt, objOpt, objOpt, true, objOpt, objOpt, true, objOpt, objOpt, objOpt, objOpt, objOpt);
                // TODO office 2003
                // excelApp.ActiveWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, (object)outfile, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt, objOpt);
            }
            catch (Exception ex)
            {
                log.ErrorFormat("Excel {0} 转换PDF出错，异常信息: {1}", infile, ex.Message);
            }
            finally
            {
                if (excelApp != null)
                {
                    excelApp.Quit();
                    excelApp = null;
                }
                    
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        /// <summary>
        /// WORD文檔轉成PDF文檔
        /// 參考 http://msdn.microsoft.com/en-us/library/bb256835%28v=office.12%29.aspx        
        /// 
        ///
        /// </summary>
        /// <param name="infile"></param>
        /// <param name="outfile"></param>
        public void Word2PDF(string infile, string outfile)
        {
            // TODO office 2013

            /*
            object objOpt = Missing.Value;
            object readOnly = true;
            object missing=Missing.Value;
            object file=(object)infile;
            object SavePDFFormat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF;

            Word.WdExportFormat paramExportFormat = Word.WdExportFormat.wdExportFormatPDF;
            bool paramOpenAfterExport = false;
            Word.WdExportOptimizeFor paramExportOptimizeFor =
            Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForPrint;
            Word.WdExportRange paramExportRange = Word.WdExportRange.wdExportAllDocument;
            int paramStartPage = 0;
            int paramEndPage = 0;
            Word.WdExportItem paramExportItem = Word.WdExportItem.wdExportDocumentContent;
            bool paramIncludeDocProps = true;
            bool paramKeepIRM = true;
            Word.WdExportCreateBookmarks paramCreateBookmarks =
            Word.WdExportCreateBookmarks.wdExportCreateWordBookmarks;
            bool paramDocStructureTags = true;
            bool paramBitmapMissingFonts = true;
            bool paramUseISO19005_1 = false;

            Word.Application wordApp = null;
            try
            {
                wordApp = new Word.Application();
                wordApp.Documents.Open(ref file, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                wordApp.ActiveDocument.ExportAsFixedFormat(outfile, paramExportFormat, paramOpenAfterExport, paramExportOptimizeFor, paramExportRange, paramStartPage,
                            paramEndPage, paramExportItem, paramIncludeDocProps,
                            paramKeepIRM, paramCreateBookmarks, paramDocStructureTags,
                            paramBitmapMissingFonts, paramUseISO19005_1,
                            ref missing);
            }
            catch (Exception ex)
            {
                log.ErrorFormat("Word {0} 转换PDF出错，异常信息: {1}", infile, ex.Message);
            }
            finally
            {
                if (wordApp != null)
                {
                    wordApp.Quit(ref missing, ref missing, ref missing);
                    wordApp = null;
                }
                    
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            */
        }


        public void PPT2PDF(string infile, string outfile)
        {
            // TODO office 2013
            // PpSaveAsFileType targetFileType = PpSaveAsFileType.ppSaveAsPDF;
            // PPTAs(infile, outfile, targetFileType);
           
        }

        public void PPT2Image(string infile, string outfile)
        {
            PpSaveAsFileType targetFileType = PpSaveAsFileType.ppSaveAsPNG;
            PPTAs(infile, outfile, targetFileType);

        }
        
        private void PPTAs(string infile, string outfile, PpSaveAsFileType targetFileType)
        {
            object missing = Type.Missing;
            Presentation persentation = null;
            PowerPoint.Application pptApp = null;
            try
            {
                pptApp = new PowerPoint.Application();
                persentation = pptApp.Presentations.Open(infile, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                persentation.SaveAs(outfile, targetFileType, Microsoft.Office.Core.MsoTriState.msoTrue);
            }
            catch (Exception ex)
            {
                log.ErrorFormat("PPT {0} 转换出错，异常信息: {1}", infile, ex.Message);
            }
            finally
            {
                if (persentation != null)
                {
                    persentation.Close();
                    persentation = null;
                }
                if (pptApp != null)
                {
                    pptApp.Quit();
                    pptApp = null;
                }
               
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void SplitExcel(string filePath, string outDir)
        {
            object oMissing = System.Reflection.Missing.Value;
            Excel.Application excel = new Excel.Application();
            Workbook workbook = excel.Workbooks.Open(filePath, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            if (File.Exists(filePath))
            {
                try
                {
                    if(!Directory.Exists(outDir))
                    {
                        Directory.CreateDirectory(outDir);
                    }
                    int i = 0;
                    string dest = filePath.Substring(0, filePath.LastIndexOf('.'));
                    foreach (Worksheet s in workbook.Worksheets)
                    {
                        i++;
                        Workbook wb = excel.Workbooks.Add(true);
                        Worksheet sheet = (Worksheet)wb.ActiveSheet;
                        s.Copy(sheet, Type.Missing);
                        sheet.Delete();
                        wb.SaveCopyAs(outDir + "\\" + i);
                        wb.Close(false, Type.Missing, Type.Missing);
                    }
                    workbook.Close(false, Type.Missing, Type.Missing);
                    excel.Quit();
                }
                catch (Exception ex)
                {
                    log.ErrorFormat("Excel {0} 分割出错，异常信息: {1}", filePath, ex.Message);
                }
            }
        }

    }
}
