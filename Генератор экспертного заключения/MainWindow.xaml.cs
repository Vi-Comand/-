using Microsoft.Office.Interop.Word;
using pdfforge.PDFCreator.UI.ComWrapper;
using System;
using System.Collections.Generic;
using System.Drawing;
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

        string p1;
        string p2;
        string p3;
        string p4;
        string p5;
        private readonly string _startPathFolder = Directory.GetCurrentDirectory();
        private Word.Application _wordapp;
        private Word.Document _worddocument;
        winForms.FolderBrowserDialog folderBrowserDialog1 = new winForms.FolderBrowserDialog();
        winForms.OpenFileDialog fileDialog1 = new winForms.OpenFileDialog();
        Word.Paragraph wordParag;
        Word.Table wordTable;
        Excel.Application objWorkExcel = new Excel.Application();



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
            Object template2 = _startPathFolder + "\\Шаблан_акт.docx";
            Object template3 = _startPathFolder + "\\Шаблон_фото.docx";
            if (FIO != "" || auto != "" || gosN != "")
            {
                string path = _startPathFolder + "\\" + FIO;
                DirectoryInfo dirInfo = new DirectoryInfo(path);
                dirInfo.Create();

                _worddocument = _wordapp.Documents.Add(ref template, ref newTemplate, ref documentType);
                _worddocument.SaveAs(path + "\\" + auto + " " + gosN + " " + FIO + ".doc");
                _worddocument.Close();

                _wordapp.Documents.Open(path + "\\" + auto + " " + gosN + " " + FIO + ".doc");
                _worddocument = _wordapp.ActiveDocument;
                try
                {
                    string filename = p3;
                    Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(filename);
                    Excel.Worksheet objWorkSheet = (Excel.Worksheet)objWorkBook.Sheets[1];
                    Excel.Range range = objWorkSheet.get_Range("B:B").Find("");
                    var excelcells = objWorkSheet.get_Range("A1", "D20").Copy();
                    wordTable = _worddocument.Tables[5];
                    cellRange = wordTable.Cell(1, 1).Range;
                    cellRange.Paste();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                //    _wordapp.Selection.InsertFile(_pathFolder + "\\" + var + ".doc", Type.Missing, false);

                int col_v_doc3 = 1;
                wordTable = _worddocument.Tables[6];
                for (int i = 0; i < l3.Items.Count; i++)
                {
                    var format = l3.Items[i].ToString().Remove(0, l3.Items[i].ToString().Length - 3);

                    if (format != "pdf")
                    {
                        cellRange = wordTable.Cell(col_v_doc3, 1).Range;
                        var shape = cellRange.InlineShapes.AddPicture(l3.Items[i].ToString(), Type.Missing, Type.Missing, Type.Missing);
                        float w = shape.Width;
                        float h = shape.Height;
                        float c = w / 950;
                        h = h / c;
                        if (h < 550)
                        {
                            shape.Width = 950;
                            shape.Height = h;
                        }
                        else
                        {
                            h = shape.Height;
                            c = h / 550;
                            w = w / c;
                            shape.Width = w;
                            shape.Height = 550;
                        }
                        wordTable.Rows.Add();
                        col_v_doc3++;
                    }
                    else
                    {
                        var names = ConvertPDFtoHojas(l3.Items[i].ToString(), path);
                        foreach (var name in names)
                        {
                            cellRange = wordTable.Cell(col_v_doc3, 1).Range;
                            var shape = cellRange.InlineShapes.AddPicture(path + @"\" + name, Type.Missing, Type.Missing, Type.Missing);
                            float w = shape.Width;
                            float h = shape.Height;
                            //float c = w / 950;
                            //h = h / c;
                            //if (h < 550)
                            //{
                            //    shape.Width = 950;
                            //    shape.Height = h;
                            //}
                            //else
                            //{
                            //    h = shape.Height;
                            //    c = h / 550;
                            //    w = w / c;
                            //    shape.Width = w;
                            //    shape.Height = 550;
                            //}

                            wordTable.Rows.Add();
                            col_v_doc3++;
                            System.IO.File.Delete(path + @"\" + name);

                        }

                    }
                }


                int col_v_doc4 = 1;
                wordTable = _worddocument.Tables[7];
                for (int i = 0; i < l4.Items.Count; i++)
                {
                    var format = l4.Items[i].ToString().Remove(0, l4.Items[i].ToString().Length - 3);

                    if (format != "pdf")
                    {
                        cellRange = wordTable.Cell(col_v_doc3, 1).Range;
                        var shape = cellRange.InlineShapes.AddPicture(l4.Items[i].ToString(), Type.Missing, Type.Missing, Type.Missing);
                        float w = shape.Width;
                        float h = shape.Height;
                        float c = w / 950;
                        h = h / c;
                        if (h < 550)
                        {
                            shape.Width = 950;
                            shape.Height = h;
                        }
                        else
                        {
                            h = shape.Height;
                            c = h / 550;
                            w = w / c;
                            shape.Width = w;
                            shape.Height = 550;
                        }
                        wordTable.Rows.Add();
                        col_v_doc4++;
                    }
                    else
                    {
                        var names = ConvertPDFtoHojas(l4.Items[i].ToString(), path);
                        foreach (var name in names)
                        {
                            cellRange = wordTable.Cell(col_v_doc3, 1).Range;
                            var shape = cellRange.InlineShapes.AddPicture(path + @"\" + name, Type.Missing, Type.Missing, Type.Missing);
                            float w = shape.Width;
                            float h = shape.Height;
                            //float c = w / 950;
                            //h = h / c;
                            //if (h < 550)
                            //{
                            //    shape.Width = 950;
                            //    shape.Height = h;
                            //}
                            //else
                            //{
                            //    h = shape.Height;
                            //    c = h / 550;
                            //    w = w / c;
                            //    shape.Width = w;
                            //    shape.Height = 550;
                            //}

                            wordTable.Rows.Add();
                            col_v_doc4++;
                            System.IO.File.Delete(path + @"\" + name);

                        }

                    }
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

                if (chec.IsChecked == false)
                {
                    _worddocument.SaveAs2(path + "\\" + auto + " " + gosN + " " + FIO + ".pdf", WdSaveFormat.wdFormatPDF);
                    _worddocument.Save();
                    _worddocument.Close();

                    _worddocument = _wordapp.Documents.Add(ref template2, ref newTemplate, ref documentType);
                    _worddocument.SaveAs(path + "\\Акт осмотра" + auto + " " + gosN + " " + FIO + ".doc");
                    _worddocument.Close();
                    _worddocument = _wordapp.Documents.Add(ref template3, ref newTemplate, ref documentType);
                    _worddocument.SaveAs(path + "\\Фото" + auto + " " + gosN + " " + FIO + ".doc");
                    _worddocument.Close();

                    _wordapp.Documents.Open(path + "\\Фото" + auto + " " + gosN + " " + FIO + ".doc");
                    _worddocument = _wordapp.ActiveDocument;
                    wordTable = _worddocument.Tables[1];
                    int col_v_doc2 = 1;
                    for (int i = 1; i <= l2.Items.Count;)
                    {
                        try
                        {
                            for (int j = 1; j < 3; j++)
                            {
                                cellRange = wordTable.Cell(i, j).Range;
                                var shape = cellRange.InlineShapes.AddPicture(l2.Items[(i - 1)].ToString(), Type.Missing, Type.Missing, Type.Missing);
                                float w = shape.Width;
                                float h = shape.Height;
                                float c = w / 480;
                                h = h / c;
                                if (h < 250)
                                {
                                    shape.Width = 480;
                                    shape.Height = h;
                                }
                                else
                                {
                                    h = shape.Height;
                                    c = h / 250;
                                    w = w / c;
                                    shape.Width = w;
                                    shape.Height = 250;
                                }
                                i++;
                            }
                            wordTable.Rows.Add();
                        }
                        catch { }
                    }
                    _worddocument.SaveAs2(path + "\\Фото" + auto + " " + gosN + " " + FIO + ".pdf", WdSaveFormat.wdFormatPDF);
                    _worddocument.Save();
                    _worddocument.Close();

                    _wordapp.Documents.Open(path + "\\Акт осмотра" + auto + " " + gosN + " " + FIO + ".doc");
                    _worddocument = _wordapp.ActiveDocument;
                    wordTable = _worddocument.Tables[1];
                    int col_v_doc1 = 1;

                    for (int i = 0; i < l1.Items.Count; i++)
                    {
                        var format = l1.Items[i].ToString().Remove(0, l1.Items[i].ToString().Length - 3);

                        if (format != "pdf")
                        {
                            cellRange = wordTable.Cell(col_v_doc3, 1).Range;
                            var shape = cellRange.InlineShapes.AddPicture(l1.Items[i].ToString(), Type.Missing, Type.Missing, Type.Missing);
                            float w = shape.Width;
                            float h = shape.Height;
                            float c = w / 950;
                            h = h / c;
                            if (h < 550)
                            {
                                shape.Width = 950;
                                shape.Height = h;
                            }
                            else
                            {
                                h = shape.Height;
                                c = h / 550;
                                w = w / c;
                                shape.Width = w;
                                shape.Height = 550;
                            }
                            wordTable.Rows.Add();
                            col_v_doc1++;
                        }
                        else
                        {
                            var names = ConvertPDFtoHojas(l1.Items[i].ToString(), path);
                            foreach (var name in names)
                            {
                                cellRange = wordTable.Cell(col_v_doc3, 1).Range;
                                var shape = cellRange.InlineShapes.AddPicture(path + @"\" + name, Type.Missing, Type.Missing, Type.Missing);
                                float w = shape.Width;
                                float h = shape.Height;
                                //float c = w / 950;
                                //h = h / c;
                                //if (h < 550)
                                //{
                                //    shape.Width = 950;
                                //    shape.Height = h;
                                //}
                                //else
                                //{
                                //    h = shape.Height;
                                //    c = h / 550;
                                //    w = w / c;
                                //    shape.Width = w;
                                //    shape.Height = 550;
                                //}

                                wordTable.Rows.Add();
                                col_v_doc1++;
                                System.IO.File.Delete(path + @"\" + name);

                            }

                        }
                    }
                    _worddocument.SaveAs2(path + "\\Акт осмотра" + auto + " " + gosN + " " + FIO + ".pdf", WdSaveFormat.wdFormatPDF);
                    _worddocument.Save();
                    _worddocument.Close();

                }
                else
                {

                    wordTable = _worddocument.Tables[8];
                    int col_v_doc1 = 1;

                    for (int i = 0; i < l1.Items.Count; i++)
                    {
                        var format = l1.Items[i].ToString().Remove(0, l1.Items[i].ToString().Length - 3);

                        if (format != "pdf")
                        {
                            cellRange = wordTable.Cell(col_v_doc3, 1).Range;
                            var shape = cellRange.InlineShapes.AddPicture(l1.Items[i].ToString(), Type.Missing, Type.Missing, Type.Missing);
                            float w = shape.Width;
                            float h = shape.Height;
                            float c = w / 950;
                            h = h / c;
                            if (h < 550)
                            {
                                shape.Width = 950;
                                shape.Height = h;
                            }
                            else
                            {
                                h = shape.Height;
                                c = h / 550;
                                w = w / c;
                                shape.Width = w;
                                shape.Height = 550;
                            }
                            wordTable.Rows.Add();
                            col_v_doc1++;
                        }
                        else
                        {
                            var names = ConvertPDFtoHojas(l1.Items[i].ToString(), path);
                            foreach (var name in names)
                            {
                                cellRange = wordTable.Cell(col_v_doc3, 1).Range;
                                var shape = cellRange.InlineShapes.AddPicture(path + @"\" + name, Type.Missing, Type.Missing, Type.Missing);
                                float w = shape.Width;
                                float h = shape.Height;
                                //float c = w / 950;
                                //h = h / c;
                                //if (h < 550)
                                //{
                                //    shape.Width = 950;
                                //    shape.Height = h;
                                //}
                                //else
                                //{
                                //    h = shape.Height;
                                //    c = h / 550;
                                //    w = w / c;
                                //    shape.Width = w;
                                //    shape.Height = 550;
                                //}

                                wordTable.Rows.Add();
                                col_v_doc1++;
                                System.IO.File.Delete(path + @"\" + name);

                            }

                        }
                    }

                    var _currentRange = _worddocument.Range(_worddocument.Content.End - 1, _worddocument.Content.End);
                    _currentRange.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
                    _currentRange.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;

                    _currentRange = _worddocument.Range(_worddocument.Content.End - 1, _worddocument.Content.End);
                    _worddocument.Tables.Add(_currentRange, 1, 1);
                    wordTable = _worddocument.Tables[9];
                    int col_v_doc2 = 1;
                    for (int i = 1; i <= l2.Items.Count;)
                    {
                        try
                        {
                            for (int j = 1; j < 3; j++)
                            {
                                cellRange = wordTable.Cell(i, j).Range;
                                var shape = cellRange.InlineShapes.AddPicture(l2.Items[(i - 1)].ToString(), Type.Missing, Type.Missing, Type.Missing);
                                float w = shape.Width;
                                float h = shape.Height;
                                float c = w / 480;
                                h = h / c;
                                if (h < 250)
                                {
                                    shape.Width = 480;
                                    shape.Height = h;
                                }
                                else
                                {
                                    h = shape.Height;
                                    c = h / 250;
                                    w = w / c;
                                    shape.Width = w;
                                    shape.Height = 250;
                                }
                                i++;
                            }
                            wordTable.Rows.Add();
                        }
                        catch { }
                    }
                    _worddocument.SaveAs2(path + "\\" + auto + " " + gosN + " " + FIO + ".pdf", WdSaveFormat.wdFormatPDF);
                    _worddocument.Save();
                    _worddocument.Close();
                }



            }
            else
            {
                MessageBox.Show("Не заполнены поля");
            }
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

        public List<string> ConvertPDFtoHojas(string filename, String dirOut)
        {
            PDFLibNet32.PDFWrapper _pdfDoc = new PDFLibNet32.PDFWrapper();
            _pdfDoc.LoadPDF(filename);
            List<string> mas = new List<string>();
            for (int i = 0; i < _pdfDoc.PageCount; i++)
            {

                System.Drawing.Image img = RenderPage(_pdfDoc, i);
                string name = string.Format("{0}{1}.jpg", i, DateTime.Now.ToString("mmss"));
                img.Save(System.IO.Path.Combine(dirOut, name));
                mas.Add(name);
            }
            _pdfDoc.Dispose();
            return mas;
        }
        public System.Drawing.Image RenderPage(PDFLibNet32.PDFWrapper doc, int page)
        {
            doc.CurrentPage = page + 1;
            doc.CurrentX = 0;
            doc.CurrentY = 0;
            doc.RenderDPI = 300;
            doc.RenderPage(IntPtr.Zero);

            // create an image to draw the page into
            var buffer = new Bitmap(doc.PageWidth, doc.PageHeight);
            doc.ClientBounds = new System.Drawing.Rectangle(0, 0, doc.PageWidth, doc.PageHeight);
            using (var g = Graphics.FromImage(buffer))
            {
                var hdc = g.GetHdc();
                try
                {
                    doc.DrawPageHDC(hdc);
                }
                finally
                {
                    g.ReleaseHdc();
                }
            }
            return buffer;

        }




    }
}
