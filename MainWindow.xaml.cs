using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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

namespace ExcelTool
{
    enum TarType
    {
        ALL,
        SEL
    }

    enum GenType
    {
        NONE,
        CODE,
        DATA
    }

    enum CodeType
    {
        CLIENT,
        SERVER
    }

    class SheetArg
    {
        public string excelName;
        public string sheetName;
    }


    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public static MainWindow instance;

        Paragraph consoleLines = new Paragraph();

        static Brush normalColor = new SolidColorBrush(Colors.Gray);
        static Brush lightColor = new SolidColorBrush(Colors.Green);
        static Brush errColor = new SolidColorBrush(Colors.Red);

        TarType tarType = TarType.ALL;
        GenType genType = GenType.NONE;
        CodeType codeType = CodeType.CLIENT;
        SheetArg selectedArg = null;

        bool initialized = false;
        int successCount = 0;
        int failedCount = 0;
        int totalSheetCount = 0;

        public MainWindow()
        {
            instance = this;
            InitializeComponent();
            initialized = true;

            this.Left = (SystemParameters.PrimaryScreenWidth - this.Width) * 0.5;
            this.Top = (SystemParameters.PrimaryScreenHeight - this.Height) * 0.5;

            richTextBox.Document.Blocks.Add(consoleLines);

            progressBar.Visibility = Visibility.Hidden;

            if (SettingWindow.checkValid())
            {
                addLog("表格:" + Properties.Settings.Default.excelPath, lightColor);
                addLog("输出:" + Properties.Settings.Default.outputPath, lightColor);
                showList();
            }

            bool defaultClient = Properties.Settings.Default.defaultClient;
            if (!defaultClient)
            {
                comboBox.SelectedIndex = 1;
            }
        }

        private void settingBtn_Click(object sender, RoutedEventArgs e)
        {
            SettingWindow.open();
        }

        void addLog(string log, Brush c)
        {
            if (consoleLines.Inlines.Count > 0)
            {
                consoleLines.Inlines.Add(new Run() { Text = "\n" });
            }

            Run r = new Run()
            {
                Text = log,
                Foreground = c
            };
            consoleLines.Inlines.Add(r);

            if ((int)richTextBox.VerticalOffset >= (int)(richTextBox.ExtentHeight-richTextBox.ViewportHeight))
            {
                richTextBox.ScrollToEnd();
            }
        }

        IWorkbook loadWorkbook(string dir, string name)
        {
            if (name.StartsWith("~") || !(name.EndsWith(".xlsx") || name.EndsWith(".xls")))
            {
                return null;
            }

            IWorkbook ret = null;
            var fs = new FileStream(System.IO.Path.Combine(dir, name), FileMode.Open, FileAccess.Read);
            if (name.EndsWith(".xlsx"))
            {
                ret = new XSSFWorkbook(fs);
            }
            else if (name.EndsWith(".xls"))
            {
                ret = new HSSFWorkbook(fs);
            }

            fs.Close();
            return ret;
        }

        bool checkSheetName(string name)
        {
            if (name.StartsWith("#"))
            {
                return false;
            }
            return true;
        }

        void showList(string filter = "")
        {
            treeView.Items.Clear();

            string excelPath = Properties.Settings.Default.excelPath;
            if (!Directory.Exists(excelPath))
            {
                addLog("路径不存在:" + excelPath, errColor);
                return;
            }

            Dictionary<string, string> allSheetNames = new Dictionary<string, string>();
            DirectoryInfo dirInfo = new DirectoryInfo(excelPath);
            foreach(FileInfo f in dirInfo.GetFiles())
            {
                {
                    try
                    {
                        string filename = "[" + f.Name.Substring(0, f.Name.IndexOf(".")) + "]";
                        IWorkbook workbook = loadWorkbook(excelPath, f.Name);
                        if (workbook == null)
                        {
                            continue;
                        }

                        for (int i = 0; i < workbook.NumberOfSheets; i++)
                        {
                            string sheetName = workbook.GetSheetName(i);
                            if (!allSheetNames.ContainsKey(sheetName))
                            {
                                allSheetNames[sheetName] = f.Name.Substring(0, f.Name.IndexOf("."));
                            }
                            else
                            {
                                string msg = "<" + allSheetNames[sheetName] + "> 和 <" + f.Name.Substring(0, f.Name.IndexOf(".")) + "> 都包含 " + sheetName;
                                if (MessageBox.Show(msg, "Sheet名重复") == MessageBoxResult.OK)
                                {
                                    Close();
                                }
                                return;
                            }

                            if (checkSheetName(sheetName))
                            {
                                string header = filename + " " + workbook.GetSheetName(i);
                                if (filter != "" && header.ToLower().IndexOf(filter.ToLower()) < 0)
                                {
                                    continue;
                                }
                                TreeViewItem sheet = new TreeViewItem()
                                {
                                    Header = header,
                                    Tag = new SheetArg()
                                    {
                                        excelName = f.Name,
                                        sheetName = sheetName
                                    }
                                };

                                totalSheetCount++;
                                treeView.Items.Add(sheet);
                            }
                        }

                        workbook.Close();
                    }
                    catch (Exception e)
                    {
                        addLog(e.Message, errColor);
                    }
                }
            }
        }

        private void genCodeAll_Click(object sender, RoutedEventArgs e)
        {
            tarType = TarType.ALL;
            genType = GenType.CODE;
            start();
        }

        private void genCodeSel_Click(object sender, RoutedEventArgs e)
        {
            tarType = TarType.SEL;
            genType = GenType.CODE;
            start();
        }

        private void genDataAll_Click(object sender, RoutedEventArgs e)
        {
            tarType = TarType.ALL;
            genType = GenType.DATA;
            start();
        }

        private void genDataSel_Click(object sender, RoutedEventArgs e)
        {
            tarType = TarType.SEL;
            genType = GenType.DATA;
            start();
        }

        private void textBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            showList(textBox.Text);
        }

        private void start()
        {
            if (tarType == TarType.SEL && selectedArg == null)
            {
                MessageBox.Show("先选择一个吧", "没有选择的对象");
                return;
            }

            string outputPath = Properties.Settings.Default.outputPath;
            if (!Directory.Exists(outputPath))
            {
                MessageBox.Show("输出路径不存在");
                return;
            }

            progressBar.Visibility = Visibility.Visible;
            Task.Factory.StartNew(gen);
        }

        private void update(string msg, Brush b, int value)
        {
            addLog(msg, b);
            progressBar.Value = value;
        }

        private void finished()
        {
            progressBar.Visibility = Visibility.Hidden;
            string msg = string.Format("生成{0}完毕，失败{1}个，成功{2}个",
                                        genType == GenType.CODE ? "代码" : "数据",
                                        failedCount,
                                        successCount);
            addLog(msg, failedCount == 0 ? lightColor : errColor);
        }

        void handleSheet(ISheet sheet, string excelName)
        {
            string msg = "[" + excelName.Substring(0, excelName.IndexOf(".")) + "] " + sheet.SheetName + " : ";

            if (sheet.LastRowNum < 4)
            {
                throw new System.Exception(msg + "内容格式不符合规范");
            }

            try
            {
                if (genType == GenType.CODE)
                {
                    if (codeType == CodeType.CLIENT)
                    {
                        Generater.genClientCode(sheet);
                    }
                    else if (codeType == CodeType.SERVER)
                    {
                        addLog("TODO", normalColor);
                        //Generater.genServerCode(sheet);
                    }
                }
                else if (genType == GenType.DATA)
                {
                    if (codeType == CodeType.CLIENT)
                    {
                        addLog("TODO", normalColor);
                        //Generater.genClientData(sheet);
                    }
                    else if (codeType == CodeType.SERVER)
                    {
                        addLog("TODO", normalColor);
                        //Generater.genServerData(sheet);
                    }
                }

                successCount++;
                this.Dispatcher.BeginInvoke(new Action<string, Brush, int>(update), msg + "^-^", normalColor, successCount * 100 / totalSheetCount);
            }
            catch (Exception e)
            {
                msg = "[ERROR] " + sheet.SheetName + ":  ";
                this.Dispatcher.BeginInvoke(new Action<string, Brush>(addLog), msg, errColor);
                throw;
            }
        }

        void handleWorkbook(IWorkbook workbook, string excelName)
        {
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                string sheetName = workbook.GetSheetName(i);
                if (checkSheetName(sheetName))
                {
                    handleSheet(workbook.GetSheet(sheetName), excelName);
                }
            }
        }

        private void gen()
        {
            string excelPath = Properties.Settings.Default.excelPath;
            successCount = 0;
            failedCount = 0;

            try
            {
                if (tarType == TarType.ALL)
                {
                    DirectoryInfo dirInfo = new DirectoryInfo(excelPath);

                    foreach (FileInfo f in dirInfo.GetFiles())
                    {
                        //Thread.Sleep(100);

                        IWorkbook workbook = loadWorkbook(excelPath, f.Name);
                        if (workbook == null)
                        {
                            continue;
                        }

                        handleWorkbook(workbook, f.Name);

                        workbook.Close();

                    }
                }
                else if (tarType == TarType.SEL)
                {
                    IWorkbook workbook = loadWorkbook(excelPath, selectedArg.excelName);
                    if (workbook == null)
                    {
                        return;
                    }

                    totalSheetCount = 1;
                    handleSheet(workbook.GetSheet(selectedArg.sheetName), selectedArg.excelName);

                    workbook.Close();
                }
            }
            catch (Exception e)
            {
                failedCount++;
                this.Dispatcher.BeginInvoke(new Action<string, Brush>(addLog), e.Message, errColor);
            }
            this.Dispatcher.BeginInvoke(new Action(finished));
        }

        private void comboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            codeType = (comboBox.SelectedIndex == 0 ? CodeType.CLIENT : CodeType.SERVER);
            if (initialized)
            {
                Properties.Settings.Default.defaultClient = (codeType == CodeType.CLIENT);
                Properties.Settings.Default.Save();
            }
        }

        private void treeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (treeView.SelectedItem != null)
            {
                selectedArg = (SheetArg)(treeView.SelectedItem as TreeViewItem).Tag;
            }
            else
            {
                selectedArg = null;
            }
        }
    }
}
