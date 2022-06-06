using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelEditor
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            DialogResult result = dialog.ShowDialog();

            textBox.Text = dialog.SelectedPath;
        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog dialog2 = new FolderBrowserDialog();
            DialogResult result = dialog2.ShowDialog();

            textBox2.Text = dialog2.SelectedPath;
        }

        private async void button3_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                pbStatus.Value = 0;

                Action action = () => { pbStatus.Value++; };
                var task = new Task(() =>
                {
                    for (var i = 0; i < 40; i++)
                    {
                        pbStatus.Dispatcher.Invoke(action);
                        Thread.Sleep(100);
                    }
                });
                task.Start();

                string fileTemplate = System.IO.Path.GetTempFileName() + ".xlsx";
                File.WriteAllBytes(fileTemplate, Properties.Resources.Склад);
                Work work = new Work();
                string currentCatalog = textBox.Text;
                string pathResult = textBox2.Text;
                await work.StartProcess(fileTemplate, currentCatalog, pathResult);

                pbStatus.Value = 100;
                System.Windows.MessageBox.Show("Готово");
            }
            catch(Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
            }
           
        }
    }
}
