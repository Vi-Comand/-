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

namespace Генератор_экспертного_заключения
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
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

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {

        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {

        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {

        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {

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
        }
    }
}
