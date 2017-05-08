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
	public enum SelType
    {
        ALL,
        SEL
    }

	public enum GenType
    {
        NONE,
        CODE,
        DATA
    }

	public enum CodeType
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

        public SelType selType = SelType.ALL;
		public GenType genType = GenType.NONE;
		public CodeType codeType = CodeType.CLIENT;
        SheetArg selectedArg = null;

        bool initialized = false;
        int successCount = 0;
        int failedCount = 0;
        int totalSheetCount = 0;
		int totalClientCount = 0;
		int totalServerCount = 0;

		Dictionary<string, string> allSheetNames = new Dictionary<string, string>();


		public MainWindow()
        {
            instance = this;
            InitializeComponent();
            initialized = true;

            this.Left = (SystemParameters.PrimaryScreenWidth - this.Width) * 0.5;
            this.Top = (SystemParameters.PrimaryScreenHeight - this.Height) * 0.5;

            logRichTextBox.Document.Blocks.Add(consoleLines);

            progressBar.Visibility = Visibility.Hidden;

			//if (SettingWindow.checkValid())
			if (Properties.Settings.Default.firstRun)
			{
				Properties.Settings.Default.firstRun = false;
				Properties.Settings.Default.Save();
				SettingWindow.open();
			}
			else if (Properties.Settings.Default.excelPath != "")
			{
                addLog("表格目录:" + Properties.Settings.Default.excelPath, lightColor);
                showList();
            }
			else
			{
				addLog("表格目录没有指定", errColor);
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

        public void addLog(string log, Brush c)
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

            if ((int)logRichTextBox.VerticalOffset >= (int)(logRichTextBox.ExtentHeight-logRichTextBox.ViewportHeight))
            {
                logRichTextBox.ScrollToEnd();
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

		public void countSheet(ISheet sheet, ref bool forClient, ref bool forServer)
		{
			IRow r4 = sheet.GetRow(3);

			for (int i = 0; i < r4.LastCellNum; i++)
			{
				ICell c = r4.GetCell(i);
				c.SetCellType(CellType.String);
				if (c.StringCellValue.ToLower().Contains("c"))
				{
					forClient = true;
				}

				if (c.StringCellValue.ToLower().Contains("s"))
				{
					forServer = true;
				}
			}
		}

        public void showList(string filter = "")
        {
            treeView.Items.Clear();

            string excelPath = Properties.Settings.Default.excelPath;
            if (!Directory.Exists(excelPath))
            {
                addLog("路径不存在:" + excelPath, errColor);
                return;
            }

            DirectoryInfo dirInfo = new DirectoryInfo(excelPath);

			if (filter == "")
			{
				totalClientCount = 0;
				totalServerCount = 0;
				allSheetNames.Clear();
			}

			foreach (FileInfo f in dirInfo.GetFiles())
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
						if (checkSheetName(sheetName))
						{
							bool forClient = false;
							bool forServer = false;
							countSheet(workbook.GetSheet(sheetName), ref forClient, ref forServer);

							if ((codeType == CodeType.CLIENT && !forClient) || (codeType == CodeType.SERVER && !forServer))
							{
								continue;
							}

							if (filter == "")
							{
								if (forClient)
								{
									totalClientCount++;
								}

								if (forServer)
								{
									totalServerCount++;
								}

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
							}


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

		private bool checkPath(string path, string errmsg)
		{
			bool ret = true;
			string caption = "";

			if (path == "")
			{
				caption = errmsg + "没有设置";
				ret = false;
			}
			else if (!Directory.Exists(path))
			{
				caption = errmsg + "不存在";
				ret = false;
			}

			if (!ret && MessageBox.Show("是否现在设置?", caption, MessageBoxButton.YesNo) == MessageBoxResult.Yes)
			{
				SettingWindow.open();
			}
			return ret;
		}

        private void genCodeAll_Click(object sender, RoutedEventArgs e)
        {
			if ((codeType == CodeType.CLIENT && !checkPath(Properties.Settings.Default.ccodePath, "客户端代码目录")) ||
				(codeType == CodeType.SERVER && !checkPath(Properties.Settings.Default.scodePath, "服务器代码目录")))
			{
				return;
			}

			selType = SelType.ALL;
            genType = GenType.CODE;
            start();
        }

        private void genCodeSel_Click(object sender, RoutedEventArgs e)
        {
			if ((codeType == CodeType.CLIENT && !checkPath(Properties.Settings.Default.ccodePath, "客户端代码目录")) ||
				(codeType == CodeType.SERVER && !checkPath(Properties.Settings.Default.scodePath, "服务器代码目录")))
			{
				return;
			}

			selType = SelType.SEL;
            genType = GenType.CODE;
            start();
        }

        private void genDataAll_Click(object sender, RoutedEventArgs e)
        {
			if ((codeType == CodeType.CLIENT && !checkPath(Properties.Settings.Default.cdataPath, "客户端数据目录")) ||
				(codeType == CodeType.SERVER && !checkPath(Properties.Settings.Default.sdataPath, "服务器数据目录")))
			{
				return;
			}

			selType = SelType.ALL;
            genType = GenType.DATA;
            start();
        }

        private void genDataSel_Click(object sender, RoutedEventArgs e)
        {
			if ((codeType == CodeType.CLIENT && !checkPath(Properties.Settings.Default.cdataPath, "客户端数据目录")) ||
				(codeType == CodeType.SERVER && !checkPath(Properties.Settings.Default.sdataPath, "服务器数据目录")))
			{
				return;
			}

			selType = SelType.SEL;
            genType = GenType.DATA;
            start();
        }

        private void textBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            showList(searchTextBox.Text);
        }

        public void start(bool async = true)
        {
            if (selType == SelType.SEL && selectedArg == null)
            {
                MessageBox.Show("先选择一个表格吧", "没有选择的对象");
                return;
            }

            progressBar.Visibility = Visibility.Visible;
			progressBar.Value = 0;

			if (async)
			{
				Task.Factory.StartNew(gen);
			}
			else
			{
				gen();
			}
        }

        private void update(string msg, Brush b, int value)
        {
            addLog(msg, b);
            progressBar.Value = value;
        }

        private void finished()
        {
            progressBar.Visibility = Visibility.Hidden;
            string msg = string.Format("---------- 生成{0}完毕，失败{1}个，成功{2}个 ----------",
                                        genType == GenType.CODE ? "代码" : "数据",
                                        failedCount,
                                        successCount);
            addLog(msg, failedCount == 0 ? lightColor : errColor);
		}

		//in thread
        void handleSheet(ISheet sheet, string excelName)
        {
			if (!allSheetNames.ContainsKey(sheet.SheetName))
			{
				return;
			}

            string msg = "[" + excelName.Substring(0, excelName.IndexOf(".")) + "] " + sheet.SheetName + " : ";

            if (sheet.LastRowNum < 4)
            {
                throw new System.Exception(msg + "内容格式不符合规范：行1<备注> 行2<字段名> 行3<字段类型> 行4<CS>");
            }

            try
            {
				bool ret = false;
                if (genType == GenType.CODE)
                {
                    if (codeType == CodeType.CLIENT)
                    {
						ret = Generater.genClientCode(sheet);
                    }
                    else if (codeType == CodeType.SERVER)
                    {
						throw new Exception("TODO");
						//Generater.genServerCode(sheet);
					}
                }
                else if (genType == GenType.DATA)
                {
                    if (codeType == CodeType.CLIENT)
                    {
						ret = Generater.genClientData(sheet);
					}
                    else if (codeType == CodeType.SERVER)
                    {
						throw new Exception("TODO");
						//Generater.genServerData(sheet);
					}
                }

				if (!ret)
				{
					return;
				}
                successCount++;
                this.Dispatcher.BeginInvoke(new Action<string, Brush, int>(update), msg + "^-^", normalColor, successCount * 100 / totalSheetCount);
            }
            catch (Exception)
            {
                this.Dispatcher.BeginInvoke(new Action<string, Brush>(addLog), msg + "Error", errColor);
                throw;
            }
        }

		//in thread
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

		//in thread
		private void gen()
        {
            string excelPath = Properties.Settings.Default.excelPath;
            successCount = 0;
            failedCount = 0;
			totalSheetCount = (codeType == CodeType.CLIENT ? totalClientCount : totalServerCount);

            try
            {
                if (selType == SelType.ALL)
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
                else if (selType == SelType.SEL)
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
            CodeType newType = (comboBox.SelectedIndex == 0 ? CodeType.CLIENT : CodeType.SERVER);
            if (initialized)
            {
                Properties.Settings.Default.defaultClient = (newType == CodeType.CLIENT);
                Properties.Settings.Default.Save();

				if (codeType != newType)
				{
					codeType = newType;
					showList();
					searchTextBox.Text = "";
				}
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
