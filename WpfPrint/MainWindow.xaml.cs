using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Xps;
using System.Windows.Xps.Packaging;
using System.IO;
using System.IO.Packaging;

namespace WpfPrint
{
    #region FixedDocumentSequence结构相关
    /*
     * XpsDocument result = new XpsDocument(xps, FileAccess.Read, CompressionOption.Maximum);
                FixedDocumentSequence seq = result.GetFixedDocumentSequence();

                foreach (DocumentReference docRef in seq.References)
                {
                    FixedDocument fixDoc = docRef.GetDocument(false);
                    foreach (PageContent pageContent in fixDoc.Pages)
                    {

                        FixedPage fixPage = pageContent.GetPageRoot(false);

                        Canvas canvas = new Canvas(); //在页面上画一个大的图层

                        canvas.Width = 300;

                        canvas.Height = 100;

                        canvas.Background = Brushes.Red;

                        fixPage.Children.Add(canvas);
                    }
                }
     */
    #endregion
    
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        //不能打印文本对象，例如TextBlock（会打印空白），换成Button就可以正常打印
        /*
         关于TextBlock空白的问题描述：
        1.直接在页面写Text属性可以正常打印
        2.后台绑定或指定Text属性不能正常打印，最终打印出的是空白
        3.TextBlock换成Button后无论是绑定还是指定都可以正常打印


        直接打印UserControl、Window可以正常打印， dialog.PrintVisual(this, "测试");

        PrintVisual普遍用于打印图片之类的
        PrintDocument常用于打印自定义文档
         */
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {

                //数据加载到模板-流文档
                LoadDocumentAndRender("PrintTemplate.xaml");
                string xps = System.IO.Path.Combine(Environment.CurrentDirectory, "xps.xps");
                ConverFlowToXPS(xps);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void Border_MouseLeftButtonDown(object sender, MouseEventArgs e)
        {
            this.DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void BtnPrint_Click(object sender, RoutedEventArgs e)
        {
            PrintDialog dialog = new PrintDialog();
            if (dialog.ShowDialog() == true)
            {
                
                dialog.PrintDocument(((IDocumentPaginatorSource)doc).DocumentPaginator, "测试");
            }
        }

        FlowDocument doc;
        public void LoadDocumentAndRender(string templateName)
        {
            //加载模板
            doc = (FlowDocument)Application.LoadComponent(new Uri(templateName, UriKind.RelativeOrAbsolute));
            doc.PagePadding = new Thickness(50);
            //doc.DataContext = data;
            Render(doc);
        }
        void ConverFlowToXPS(string xpsFullName)
        {
            XpsDocument xps = new XpsDocument(xpsFullName, System.IO.FileAccess.ReadWrite);
            //create a XPS document writer that writes to the XPS document
            XpsDocumentWriter xpsWriter = XpsDocument.CreateXpsDocumentWriter(xps);
            xpsWriter.Write(((IDocumentPaginatorSource)doc).DocumentPaginator);
            docViewer.Document = xps.GetFixedDocumentSequence();
            xps.Close();
        }
        void FlowToMemoryXPS()
        {
            MemoryStream ms = new MemoryStream();//准备在内存中存储内容
            Package package = Package.Open(ms, FileMode.Create, FileAccess.ReadWrite);
            Uri DocumentUri = new Uri("pack://InMemoryDocument.xps");
            PackageStore.RemovePackage(DocumentUri);
            PackageStore.AddPackage(DocumentUri, package);
            XpsDocument xpsDocument = new XpsDocument(package, CompressionOption.Fast, DocumentUri.AbsoluteUri);
        }
        /// <summary>
        /// 根据查询数据动态创建流文本
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="data"></param>
        public void Render(FlowDocument doc)
        {
            TableRowGroup group = doc.FindName("rowsDetails") as TableRowGroup;
            Style styleCell = doc.Resources["CellStyle"] as Style;
            //模拟多页数据
            for (int i = 0; i < 50; i++)
            {
                TableRow row = new TableRow();

                TableCell cell = new TableCell(new Paragraph(new Run("123")));
                cell.Style = styleCell;
                row.Cells.Add(cell);

                cell = new TableCell(new Paragraph(new Run("测试")));
                cell.Style = styleCell;
                row.Cells.Add(cell);

                cell = new TableCell(new Paragraph(new Run("描述测试")));
                cell.Style = styleCell;
                row.Cells.Add(cell);

                group.Rows.Add(row);
            }
        }
    }
}
