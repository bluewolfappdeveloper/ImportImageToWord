using Microsoft.Win32;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ImageToWord
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly BackgroundWorker worker = new BackgroundWorker();
        List<String> PicFileName;
        String filepicpath;
        String filepicsavepath;
        double contrast;

        public MainWindow()
        {
            InitializeComponent();

            worker.DoWork += worker_DoWork;
            worker.RunWorkerCompleted += worker_RunWorkerCompleted;
          
            worker.ProgressChanged += Worker_ProgressChanged;
            worker.WorkerReportsProgress = true;
        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //processSaveWord.Value = e.ProgressPercentage;
        }

        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            // run all background tasks here

            Document document = new Document();
            Paragraph paragraph = document.AddSection().AddParagraph(); ;

            document.Sections[0].PageSetup.Margins.Left = 36f;
            document.Sections[0].PageSetup.Orientation = PageOrientation.Landscape;
            document.Sections[0].PageSetup.PageSize = PageSize.A3;

            foreach (String picname in PicFileName)
            {
               //paragraph = document.AddSection().AddParagraph();

                string path = filepicpath + @"\" + picname;
                DocPicture Pic = paragraph.AppendPicture(System.Drawing.Image.FromFile(path));
                Pic.Width = 165.6f;
                Pic.Height = 243.36f;


                Pic.Contrast = (float)contrast;
            }

            

            //Save and Launch
            document.SaveToFile(filepicsavepath + @"\" + "ImageInsert.docx", FileFormat.Docx);

            try
            {
                System.Diagnostics.Process.Start(filepicsavepath + @"\" + "ImageInsert.docx");
            }
            catch { }
            
        }

        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }

        //public void InsertToWord()
        //{
        //    Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();  
        //    //winword.ShowAnimation = false;
        //    winword.Visible = false;
        //    object missing = System.Reflection.Missing.Value;

        //    Microsoft.Office.Interop.Word.Document document =  winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);
        //    Microsoft.Office.Interop.Word.Paragraph paragraph = document.Paragraphs.Add(ref missing);

        //    document.PageSetup.LeftMargin = 0;

        //    Microsoft.Office.Interop.Word.InlineShape finalInlineShape;
        //    Microsoft.Office.Interop.Word.Range docRange;

        //    for (int i = PicFileName.Count-1; i >= 0; i--)
        //    {
        //        string picname = PicFileName.ElementAt(i);
         
        //        object oCollapseEnd = Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseEnd;
        //        docRange = document.Content;
        //        docRange.Collapse(ref oCollapseEnd);
        //        string path = filepicpath + @"\" + picname;
        //        Microsoft.Office.Interop.Word.InlineShape inline_shape = paragraph.Range.InlineShapes.AddPicture(path, ref missing, ref missing, ref missing);
                
        //        // Format the picture.
        //        Microsoft.Office.Interop.Word.Shape shape = inline_shape.ConvertToShape();
        //        shape.Width = 165.6f;
        //        shape.Height = 243.36f;

        //        finalInlineShape = shape.ConvertToInlineShape();
        //    }

        //    object filename = filepicsavepath + @"\" + "ImageInsert.docx";

        //    //MessageBox.Show(document.Saved.ToString());

        //    document.SaveAs(ref filename, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault);
        //    document.Close(ref missing, ref missing, ref missing);
        //    document = null;
        //    winword.Quit(ref missing, ref missing, ref missing);
        //    winword = null;
        //    //MessageBox.Show("Document created successfully !");
        //    try
        //    {
        //        System.Diagnostics.Process.Start(filepicsavepath + @"\" + "ImageInsert.docx");
        //    }
        //    catch { }
        //}

        private void btnImageLocal_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                System.Windows.Forms.DialogResult result = dialog.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK)
                    txtURLPic.Text = dialog.SelectedPath;
                    
            }
        }

        private void btnFileLocal_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (openFileDialog.ShowDialog() == true)
            {
                txtFileSave.Text = openFileDialog.FileName;
            }
        }

        private void btnLoadImage_Click(object sender, RoutedEventArgs e)
        {
            if (File.Exists(txtFileSave.Text) == false || Directory.Exists(txtURLPic.Text) == false)
            {
                MessageBox.Show("Vui lòng kiểm tra lại đường dẫn", "Thông Báo", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string[] lines = System.IO.File.ReadAllLines(txtFileSave.Text);
            PicFileName = new List<string>();

            foreach (String filename in lines)
            {
                if (File.Exists(txtURLPic.Text + @"\" + ChangeFileName(filename) + ".jpg"))
                {
                    PicFileName.Add(ChangeFileName(filename) + ".jpg");
                }

            }

            lstNameImage.ItemsSource = PicFileName;
        }

        private string ChangeFileName(string fileName)
        {
            while (fileName[0] == '0')
            {
                fileName = fileName.Remove(0, 1);
            }
            return fileName;
        }

        private void lstNameImage_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string selectedFileName = txtURLPic.Text + @"\" + PicFileName.ElementAt(lstNameImage.SelectedIndex);
            
            BitmapImage bitmap = new BitmapImage();
            bitmap.BeginInit();
            bitmap.UriSource = new Uri(selectedFileName);
            bitmap.EndInit();
            imgShow.Source = bitmap;

            
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            
            if (worker.IsBusy || PicFileName == null || PicFileName.Count == 0)
            {
                if (worker.IsBusy)
                    MessageBox.Show("Process is running !!!", "Thông báo", MessageBoxButton.OK);

                return;
            }
            filepicpath = txtURLPic.Text;

            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                System.Windows.Forms.DialogResult result = dialog.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    filepicsavepath = dialog.SelectedPath;
                    MessageBoxResult IsRun = MessageBox.Show("Process is ready !!! CLick OK to start", "Thông báo", MessageBoxButton.OKCancel);
                    if (IsRun == MessageBoxResult.OK) worker.RunWorkerAsync();
                }
                  

            }

           
        }


        private void Slider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if ((sender as Slider).Value % 10 != 0) (sender as Slider).Value = e.OldValue;
            tbContrast.Text = "Contrast: " + ((sender as Slider).Value - 40).ToString() + "%";
            contrast = (sender as Slider).Value - 40;
        }


    }
}
