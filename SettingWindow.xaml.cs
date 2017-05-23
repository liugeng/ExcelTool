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
		Brush norColor = new SolidColorBrush(Colors.Black);
		Brush errColor = new SolidColorBrush(Colors.Red);

        public SettingWindow()
        {
            InitializeComponent();

            this.Left = MainWindow.instance.Left + MainWindow.instance.Width * 0.5 - Width * 0.5;
            this.Top = MainWindow.instance.Top + MainWindow.instance.Height * 0.3 - Height * 0.5;

            excelPathTextBox.Text = Properties.Settings.Default.excelPath;
            ccodePathTextBox.Text = Properties.Settings.Default.ccodePath;
			cscriptPathTextBox.Text = Properties.Settings.Default.cscriptPath;
			cdataPathTextBox.Text = Properties.Settings.Default.cdataPath;
			scodePathTextBox.Text = Properties.Settings.Default.scodePath;
			sdataPathTextBox.Text = Properties.Settings.Default.sdataPath;

			checkPath(excelPathLabel, Properties.Settings.Default.excelPath);
			checkPath(ccodePathLabel, Properties.Settings.Default.ccodePath);
			checkPath(cscriptPathLabel, Properties.Settings.Default.cscriptPath);
			checkPath(cdataPathLabel, Properties.Settings.Default.cdataPath);
			checkPath(scodePathLabel, Properties.Settings.Default.scodePath);
			checkPath(sdataPathLabel, Properties.Settings.Default.sdataPath);
		}

		void checkPath(Label lab, string path)
		{
			if (path == "" || !Directory.Exists(path))
			{
				lab.Foreground = errColor;
			}
			else
			{
				lab.Foreground = norColor;
			}
		}

        public static void open()
        {
            (new SettingWindow()).ShowDialog();
        }

        //public static bool checkValid()
        //{
        //    string excelPath = Properties.Settings.Default.excelPath;
        //    if (excelPath != "" && !Directory.Exists(excelPath))
        //    {
        //        excelPath = "";
        //        Properties.Settings.Default.excelPath = "";
        //    }

        //    string ccodePath = Properties.Settings.Default.ccodePath;
        //    if (ccodePath != "" && !Directory.Exists(ccodePath))
        //    {
        //        ccodePath = "";
        //        Properties.Settings.Default.ccodePath = "";
        //    }

        //    if (excelPath == "" || ccodePath == "")
        //    {
        //        Properties.Settings.Default.Save();
        //        open();
        //        return false;
        //    }
        //    return true;
        //}

        private void excelPathBtn_Click(object sender, RoutedEventArgs e)
        {
            excelPathTextBox.Text = openSettingDialog(Properties.Settings.Default.excelPath, "选择表格目录");
        }

        private void ccodePathBtn_Click(object sender, RoutedEventArgs e)
        {
            ccodePathTextBox.Text = openSettingDialog(Properties.Settings.Default.ccodePath, "选择客户端 C++ 输出目录");
        }

		private void cscriptPathBtn_Click(object sender, RoutedEventArgs e)
		{
			cscriptPathTextBox.Text = openSettingDialog(Properties.Settings.Default.cscriptPath, "选择客户端脚本输出目录");
		}

		private void cdataPathBtn_Click(object sender, RoutedEventArgs e)
		{
			cdataPathTextBox.Text = openSettingDialog(Properties.Settings.Default.ccodePath, "选择客户端数据输出目录");
		}

		private void scodePathBtn_Click(object sender, RoutedEventArgs e)
		{
			scodePathTextBox.Text = openSettingDialog(Properties.Settings.Default.ccodePath, "选择服务器代码输出目录");
		}

		private void sdataPathBtn_Click(object sender, RoutedEventArgs e)
		{
			sdataPathTextBox.Text = openSettingDialog(Properties.Settings.Default.ccodePath, "选择服务器数据输出目录");
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
            //if (!Directory.Exists(excelPathTextBox.Text))
            //{
            //    if (MessageBox.Show("表格路径不存在，是否重新设置？", "路径不正确") == MessageBoxResult.Yes)
            //    {
            //        e.Cancel = true;
            //    }
            //    else
            //    {
            //        Application.Current.Shutdown(0);
            //    }
            //    return;
            //}

            //if (!Directory.Exists(ccodePathTextBox.Text))
            //{
            //    if (MessageBox.Show("输出路径不存在，是否重新设置？", "路径不正确", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            //    {
            //        e.Cancel = true;
            //    }
            //    else
            //    {
            //        Application.Current.Shutdown(0);
            //    }
            //    return;
            //}

            Properties.Settings.Default.excelPath = excelPathTextBox.Text;
            Properties.Settings.Default.ccodePath = ccodePathTextBox.Text;
			Properties.Settings.Default.cscriptPath = cscriptPathTextBox.Text;
			Properties.Settings.Default.cdataPath = cdataPathTextBox.Text;
			Properties.Settings.Default.scodePath = scodePathTextBox.Text;
			Properties.Settings.Default.sdataPath = sdataPathTextBox.Text;
			Properties.Settings.Default.Save();
        }

		private void excelPathTextBox_TextChanged(object sender, TextChangedEventArgs e)
		{
			string path = excelPathTextBox.Text;
			checkPath(excelPathLabel, path);
			if (path != "" && Directory.Exists(path) && path != Properties.Settings.Default.excelPath)
			{
				Properties.Settings.Default.excelPath = path;
				Properties.Settings.Default.Save();
				MainWindow.instance.showList();
			}
		}

		private void ccodePathTextBox_TextChanged(object sender, TextChangedEventArgs e)
		{
			checkPath(ccodePathLabel, ccodePathTextBox.Text);
		}

		private void cscriptPathTextBox_TextChanged(object sender, TextChangedEventArgs e)
		{
			checkPath(cscriptPathLabel, cscriptPathTextBox.Text);
		}

		private void cdataPathTextBox_TextChanged(object sender, TextChangedEventArgs e)
		{
			checkPath(cdataPathLabel, cdataPathTextBox.Text);
		}

		private void scodePathTextBox_TextChanged(object sender, TextChangedEventArgs e)
		{
			checkPath(scodePathLabel, scodePathTextBox.Text);
		}

		private void sdataPathTextBox_TextChanged(object sender, TextChangedEventArgs e)
		{
			checkPath(sdataPathLabel, sdataPathTextBox.Text);
		}

	}
}
