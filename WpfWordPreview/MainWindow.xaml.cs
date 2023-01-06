using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.IO;
using System.IO.Packaging;
using System.Windows.Xps.Packaging;
using System.Runtime.InteropServices;

using Word = Microsoft.Office.Interop.Word;
using System.Windows.Xps;

namespace WpfWordPreview
{
    /// <summary>
    /// word转xps
    /// 1.无论是.net6.0还是.net4.7转换都没问题（word to xps）
    /// 2.唯一区别就是6.0可以正常显示，4.7显示不完整并且报错：StoryFragments 部分加载失败。
    /// 3.6.0只能打开以打印形式生成的xps文件，无法打开以打印形式生成的oxps文件
    /// 
    /// 
    /// 方案1
    /// 手动创建一个FixedDocumentSequence，然后将word每页转换成图片，以图片形式添加进去，同时添加logo，之后DocumentViewer展示
    /// 详情见-FixedDocumentSequence结构相关
    /// 
    /// 方案2
    /// 将word直接转为pdf，使用第三方组件展示即可
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private void Border_MouseLeftButtonDown(object sender, MouseEventArgs e)
        {
            this.DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void BtnShowXPS_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                /*
                 * 此处代码在.net6.0可以正常运行，.net4.7报错-StoryFragments 部分加载失败。
                 */
                string word = Environment.CurrentDirectory + "\\word.docx";
                string xps = Environment.CurrentDirectory + "\\word.xps";
                ConvertWordToXPS(word,xps);
                XpsDocument doc = new XpsDocument(xps, System.IO.FileAccess.Read);
                docViewer.Document= doc.GetFixedDocumentSequence();
                doc.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void BtnShowPDF_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string word = Environment.CurrentDirectory + "\\word.docx";
                string pdf = Environment.CurrentDirectory + "\\word.pdf";
                ConvertWordToPDF(word, pdf);
                moonPdfPanel.OpenFile(pdf);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void ConvertWordToXPS(string wordName, string xpsFullName)
        {
            Microsoft.Office.Interop.Word.Application wordApplication = new Microsoft.Office.Interop.Word.Application();
            try
            {
                System.IO.File.Delete(xpsFullName);
                wordApplication.Documents.Add(wordName);
                Microsoft.Office.Interop.Word.Document doc = wordApplication.ActiveDocument;
                doc.ExportAsFixedFormat(xpsFullName, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatXPS, false,
                    Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                    Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument, 0, 0,
                    Microsoft.Office.Interop.Word.WdExportItem.wdExportDocumentContent, true, true,
                    Microsoft.Office.Interop.Word.WdExportCreateBookmarks.wdExportCreateHeadingBookmarks, true, true, false, Type.Missing);
            }
            catch (Exception ex)
            {
                string error = ex.Message;
                MessageBox.Show(ex.Message);
            }
            finally
            {
                wordApplication.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
            }
        }
        private void ConvertWordToPDF(string wordName, string pdfFullName)
        {
            Microsoft.Office.Interop.Word.Application wordApplication = new Microsoft.Office.Interop.Word.Application();
            try
            {
                System.IO.File.Delete(pdfFullName);
                wordApplication.Documents.Add(wordName);
                Microsoft.Office.Interop.Word.Document doc = wordApplication.ActiveDocument;
                doc.ExportAsFixedFormat(pdfFullName, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, false,
                    Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                    Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument, 0, 0,
                    Microsoft.Office.Interop.Word.WdExportItem.wdExportDocumentContent, true, true,
                    Microsoft.Office.Interop.Word.WdExportCreateBookmarks.wdExportCreateHeadingBookmarks, true, true, false, Type.Missing);
            }
            catch (Exception ex)
            {
                string error = ex.Message;
                MessageBox.Show(ex.Message);
            }
            finally
            {
                wordApplication.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
            }
        }
    }
}
