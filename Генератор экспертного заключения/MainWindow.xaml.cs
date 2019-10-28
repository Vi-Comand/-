using Microsoft.Office.Interop.Word;
using pdfforge.PDFCreator.UI.ComWrapper;
using System;
using System.Collections.Generic;
using System.IO;
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
using Excel = Microsoft.Office.Interop.Excel;
using winForms = System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Генератор_экспертного_заключения
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
            _wordapp = new Word.Application();
        }

        private object _extend;
        private string _pathFolder;
        private readonly string _startPathFolder = Directory.GetCurrentDirectory();
        private Word.Application _wordapp;
        private Word.Document _worddocument;
        winForms.FolderBrowserDialog folderBrowserDialog1 = new winForms.FolderBrowserDialog();
        winForms.OpenFileDialog fileDialog1 = new winForms.OpenFileDialog();
        Word.Paragraph wordParag;
        Word.Table wordTable;
        string p1;
        string p2;
        string p3;
        string p4;
        string p5;

        string zak = "";
        string date = "";
        string exp = "";
        string nexp = "";
        string d = "";
        string m = "";
        string y = "";
        string auto = "";
        string godM = "";
        string gosN = "";
        string nomK = "";
        string vin = "";
        string color = "";
        string probeg = "";
        string TP = "";
        string FIO = "";
        string dFIO = "";
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (fileDialog1.ShowDialog() == winForms.DialogResult.Cancel)
            {
                return;
            }

            p1 = fileDialog1.FileName;

            l1.Items.Add(p1);
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (fileDialog1.ShowDialog() == winForms.DialogResult.Cancel)
            {
                return;
            }

            p2 = fileDialog1.FileName;

            l2.Items.Add(p2);
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            if (fileDialog1.ShowDialog() == winForms.DialogResult.Cancel)
            {
                return;
            }

            p3 = fileDialog1.FileName;
            lab1.Content = "Загружено";
            //l1.Items.Add(p1);
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            if (fileDialog1.ShowDialog() == winForms.DialogResult.Cancel)
            {
                return;
            }

            p4 = fileDialog1.FileName;

            l3.Items.Add(p4);
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            if (fileDialog1.ShowDialog() == winForms.DialogResult.Cancel)
            {
                return;
            }

            p5 = fileDialog1.FileName;

            l4.Items.Add(p5);
        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            zak = textBox1.Text;
            try { date = date1.SelectedDate.Value.Date.ToString(); } catch { }
            exp = textBox12.Text;
            nexp = textBox13.Text;
            try
            {
                d = date2.SelectedDate.Value.Day.ToString();
                m = date2.SelectedDate.Value.Month.ToString();
                y = date2.SelectedDate.Value.Year.ToString();
            }
            catch { }
            auto = textBox2.Text;
            godM = textBox3.Text;
            gosN = textBox4.Text;
            nomK = textBox5.Text;
            vin = textBox6.Text;
            color = textBox7.Text;
            probeg = textBox8.Text;
            TP = textBox9.Text;
            FIO = textBox10.Text;
            dFIO = textBox11.Text;

            Range cellRange;
            Object newTemplate = false;
            Object documentType = Word.WdNewDocumentType.wdNewBlankDocument;
            Object template = _startPathFolder + "\\Шаблон.docx";

            string path = _startPathFolder + "\\" + FIO;
            DirectoryInfo dirInfo = new DirectoryInfo(path);
            dirInfo.Create();
            _worddocument = _wordapp.Documents.Add(ref template, ref newTemplate, ref documentType);
            _worddocument.SaveAs(path + "\\" + auto + " " + gosN + " " + FIO + ".doc");
            _worddocument.Close();


            _wordapp.Documents.Open(path + "\\" + auto + " " + gosN + " " + FIO + ".doc");

            //    _wordapp.Selection.InsertFile(_pathFolder + "\\" + var + ".doc", Type.Missing, false);
            _worddocument = _wordapp.ActiveDocument;
            wordTable = _worddocument.Tables[5];

            for (int i = 1; i <= l1.Items.Count; i++)
            {
                cellRange = wordTable.Cell(i, 1).Range;
                var shape = cellRange.InlineShapes.AddPicture(l1.Items[i - 1].ToString(), Type.Missing, Type.Missing, Type.Missing);
                shape.Width = 566;
                shape.Height = 377;
                wordTable.Rows.Add();
            }
            wordTable = _worddocument.Tables[6];
            for (int i = 1; i <= l2.Items.Count; i++)
            {
                cellRange = wordTable.Cell(i, 1).Range;
                var shape = cellRange.InlineShapes.AddPicture(l2.Items[i - 1].ToString(), Type.Missing, Type.Missing, Type.Missing);
                shape.Width = 300;
                shape.Height = 150;
                wordTable.Rows.Add();
            }
            wordTable = _worddocument.Tables[7];
            for (int i = 1; i <= l3.Items.Count; i++)
            {
                cellRange = wordTable.Cell(i, 1).Range;
                var shape = cellRange.InlineShapes.AddPicture(l3.Items[i - 1].ToString(), Type.Missing, Type.Missing, Type.Missing);
                shape.Width = 566;
                shape.Height = 377;
                wordTable.Rows.Add();
            }
            wordTable = _worddocument.Tables[8];
            for (int i = 1; i <= l4.Items.Count; i++)
            {
                cellRange = wordTable.Cell(i, 1).Range;
                var shape = cellRange.InlineShapes.AddPicture(l4.Items[i - 1].ToString(), Type.Missing, Type.Missing, Type.Missing);
                shape.Width = 566;
                shape.Height = 377;
                wordTable.Rows.Add();
            }

            Object missing = Type.Missing;
            Word.Find find = _wordapp.Selection.Find;


            //    worddocument.Activate();
            find.Text = "<zak>"; // текст поиска
            find.Replacement.Text = zak; // текст замены
            find.Execute(FindText: Type.Missing, MatchCase: false, MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: missing, MatchAllWordForms: false, Forward: true,
                Wrap: Word.WdFindWrap.wdFindContinue,
                Format: false, ReplaceWith: missing, Replace: Word.WdReplace.wdReplaceAll);
            find.Text = "<date>"; // текст поиска
            find.Replacement.Text = date; // текст замены
            find.Execute(FindText: Type.Missing, MatchCase: false, MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: missing, MatchAllWordForms: false, Forward: true,
                Wrap: Word.WdFindWrap.wdFindContinue,
                Format: false, ReplaceWith: missing, Replace: Word.WdReplace.wdReplaceAll);
            find.Text = "<exp>"; // текст поиска
            find.Replacement.Text = exp; // текст замены
            find.Execute(FindText: Type.Missing, MatchCase: false, MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: missing, MatchAllWordForms: false, Forward: true,
                Wrap: Word.WdFindWrap.wdFindContinue,
                Format: false, ReplaceWith: missing, Replace: Word.WdReplace.wdReplaceAll);
            find.Text = "<nexp>"; // текст поиска
            find.Replacement.Text = nexp; // текст замены
            find.Execute(FindText: Type.Missing, MatchCase: false, MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: missing, MatchAllWordForms: false, Forward: true,
                Wrap: Word.WdFindWrap.wdFindContinue,
                Format: false, ReplaceWith: missing, Replace: Word.WdReplace.wdReplaceAll);
            find.Text = "<d>"; // текст поиска
            find.Replacement.Text = d; // текст замены
            find.Execute(FindText: Type.Missing, MatchCase: false, MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: missing, MatchAllWordForms: false, Forward: true,
                Wrap: Word.WdFindWrap.wdFindContinue,
                Format: false, ReplaceWith: missing, Replace: Word.WdReplace.wdReplaceAll);
            find.Text = "<m>"; // текст поиска
            find.Replacement.Text = m; // текст замены
            find.Execute(FindText: Type.Missing, MatchCase: false, MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: missing, MatchAllWordForms: false, Forward: true,
                Wrap: Word.WdFindWrap.wdFindContinue,
                Format: false, ReplaceWith: missing, Replace: Word.WdReplace.wdReplaceAll);
            find.Text = "<y>"; // текст поиска
            find.Replacement.Text = y; // текст замены
            find.Execute(FindText: Type.Missing, MatchCase: false, MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: missing, MatchAllWordForms: false, Forward: true,
                Wrap: Word.WdFindWrap.wdFindContinue,
                Format: false, ReplaceWith: missing, Replace: Word.WdReplace.wdReplaceAll);
            find.Text = "<auto>"; // текст поиска
            find.Replacement.Text = auto; // текст замены
            find.Execute(FindText: Type.Missing, MatchCase: false, MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: missing, MatchAllWordForms: false, Forward: true,
                Wrap: Word.WdFindWrap.wdFindContinue,
                Format: false, ReplaceWith: missing, Replace: Word.WdReplace.wdReplaceAll);
            find.Text = "<godM>"; // текст поиска
            find.Replacement.Text = godM; // текст замены
            find.Execute(FindText: Type.Missing, MatchCase: false, MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: missing, MatchAllWordForms: false, Forward: true,
                Wrap: Word.WdFindWrap.wdFindContinue,
                Format: false, ReplaceWith: missing, Replace: Word.WdReplace.wdReplaceAll);
            find.Text = "<gosN>"; // текст поиска
            find.Replacement.Text = gosN; // текст замены
            find.Execute(FindText: Type.Missing, MatchCase: false, MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: missing, MatchAllWordForms: false, Forward: true,
                Wrap: Word.WdFindWrap.wdFindContinue,
                Format: false, ReplaceWith: missing, Replace: Word.WdReplace.wdReplaceAll);
            find.Text = "<nomK>"; // текст поиска
            find.Replacement.Text = nomK; // текст замены
            find.Execute(FindText: Type.Missing, MatchCase: false, MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: missing, MatchAllWordForms: false, Forward: true,
                Wrap: Word.WdFindWrap.wdFindContinue,
                Format: false, ReplaceWith: missing, Replace: Word.WdReplace.wdReplaceAll);
            find.Text = "<vin>"; // текст поиска
            find.Replacement.Text = vin; // текст замены
            find.Execute(FindText: Type.Missing, MatchCase: false, MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: missing, MatchAllWordForms: false, Forward: true,
                Wrap: Word.WdFindWrap.wdFindContinue,
                Format: false, ReplaceWith: missing, Replace: Word.WdReplace.wdReplaceAll);
            find.Text = "<color>"; // текст поиска
            find.Replacement.Text = color; // текст замены
            find.Execute(FindText: Type.Missing, MatchCase: false, MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: missing, MatchAllWordForms: false, Forward: true,
                Wrap: Word.WdFindWrap.wdFindContinue,
                Format: false, ReplaceWith: missing, Replace: Word.WdReplace.wdReplaceAll);
            find.Text = "<probeg>"; // текст поиска
            find.Replacement.Text = probeg; // текст замены
            find.Execute(FindText: Type.Missing, MatchCase: false, MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: missing, MatchAllWordForms: false, Forward: true,
                Wrap: Word.WdFindWrap.wdFindContinue,
                Format: false, ReplaceWith: missing, Replace: Word.WdReplace.wdReplaceAll);
            find.Text = "<TP>"; // текст поиска
            find.Replacement.Text = TP; // текст замены
            find.Execute(FindText: Type.Missing, MatchCase: false, MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: missing, MatchAllWordForms: false, Forward: true,
                Wrap: Word.WdFindWrap.wdFindContinue,
                Format: false, ReplaceWith: missing, Replace: Word.WdReplace.wdReplaceAll);
            find.Text = "<FIO>"; // текст поиска
            find.Replacement.Text = FIO; // текст замены
            find.Execute(FindText: Type.Missing, MatchCase: false, MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: missing, MatchAllWordForms: false, Forward: true,
                Wrap: Word.WdFindWrap.wdFindContinue,
                Format: false, ReplaceWith: missing, Replace: Word.WdReplace.wdReplaceAll);
            find.Text = "<dFIO>"; // текст поиска
            find.Replacement.Text = dFIO; // текст замены
            find.Execute(FindText: Type.Missing, MatchCase: false, MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: missing, MatchAllWordForms: false, Forward: true,
                Wrap: Word.WdFindWrap.wdFindContinue,
                Format: false, ReplaceWith: missing, Replace: Word.WdReplace.wdReplaceAll);

            _wordapp.ActiveDocument.Save();
            _wordapp.ActiveDocument.Close();

        }

        private void och1(object sender, RoutedEventArgs e)
        {
            l1.Items.Clear();
        }

        private void och2(object sender, RoutedEventArgs e)
        {
            l2.Items.Clear();
        }

        private void och3(object sender, RoutedEventArgs e)
        {
            l3.Items.Clear();
        }

        private void och4(object sender, RoutedEventArgs e)
        {
            l4.Items.Clear();
        }
    }
}
