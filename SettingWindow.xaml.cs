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
using System.Windows.Shapes;

namespace ExcelTool
{
    /// <summary>
    /// SettingWindow.xaml 的交互逻辑
    /// </summary>
    public partial class SettingWindow : Window
    {
        public SettingWindow()
        {
            InitializeComponent();

            this.Left = MainWindow.instance.Left + MainWindow.instance.Width * 0.5 - Width * 0.5;
            this.Top = MainWindow.instance.Top + MainWindow.instance.Height * 0.3 - Height * 0.5;

            excelPathTextBox.Text = Properties.Settings.Default.excelPath;
            outputPathTextBox.Text = Properties.Settings.Default.outputPath;

            
        }

        public static void open()
        {
            (new SettingWindow()).ShowDialog();
        }

        public static bool checkValid()
        {
            string excelPath = Properties.Settings.Default.excelPath;
            if (excelPath != "" && !Directory.Exists(excelPath))
            {
                excelPath = "";
                Properties.Settings.Default.excelPath = "";
            }

            string outputPath = Properties.Settings.Default.outputPath;
            if (outputPath != "" && !Directory.Exists(outputPath))
            {
                outputPath = "";
                Properties.Settings.Default.outputPath = "";
            }

            if (excelPath == "" || outputPath == "")
            {
                Properties.Settings.Default.Save();
                open();
                return false;
            }
            return true;
        }

        private void excelPathTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void outputPathTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void excelPathBtn_Click(object sender, RoutedEventArgs e)
        {
            excelPathTextBox.Text = openSettingDialog(Properties.Settings.Default.excelPath, "选择表格目录");
        }

        private void outputPathBtn_Click(object sender, RoutedEventArgs e)
        {
            outputPathTextBox.Text = openSettingDialog(Properties.Settings.Default.outputPath, "选择输出目录");
        }

        private string openSettingDialog(String selectedPath, string des)
        {
            if (selectedPath == "")
            {
                selectedPath = Directory.GetCurrentDirectory();
            }

            System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            folderBrowserDialog.Description = des;
            folderBrowserDialog.SelectedPath = selectedPath;
            System.Windows.Forms.DialogResult result = folderBrowserDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                if (folderBrowserDialog.SelectedPath != selectedPath)
                {
                    return folderBrowserDialog.SelectedPath;
                }
            }
            return selectedPath;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            string excelPath = Properties.Settings.Default.excelPath;
            string outputPath = Properties.Settings.Default.outputPath;
            if (!Directory.Exists(excelPathTextBox.Text))
            {
                if (MessageBox.Show("表格路径不存在，是否重新设置？", "路径不正确") == MessageBoxResult.Yes)
                {
                    e.Cancel = true;
                }
                else
                {
                    Application.Current.Shutdown(0);
                }
                return;
            }

            if (!Directory.Exists(outputPathTextBox.Text))
            {
                if (MessageBox.Show("输出路径不存在，是否重新设置？", "路径不正确", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    e.Cancel = true;
                }
                else
                {
                    Application.Current.Shutdown(0);
                }
                return;
            }

            Properties.Settings.Default.excelPath = excelPathTextBox.Text;
            Properties.Settings.Default.outputPath = outputPathTextBox.Text;
            Properties.Settings.Default.Save();
        }
    }
}
