using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelTool
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : Application
    {
		protected override void OnStartup(StartupEventArgs e)
		{
			// >ExcelTool.exe -h | more
			// 命令行模式
			if (e.Args.Length > 0)
			{
				MainWindow w = new MainWindow();
				foreach (string arg in e.Args)
				{
					switch (arg)
					{
						case "-h":
							{
								showHelp();
								break;
							}
						case "-cc":
							{
								Console.WriteLine("生成客户端代码");
								w.codeType = CodeType.CLIENT;
								w.selType = SelType.ALL;
								w.genType = GenType.CODE;
								w.start(false);
								Thread.Sleep(100);
								break;
							}
						case "-cd":
							{
								Console.WriteLine("生成客户端数据");
								w.codeType = CodeType.CLIENT;
								w.selType = SelType.ALL;
								w.genType = GenType.DATA;
								w.start(false);
								Thread.Sleep(100);
								break;
							}
						case "-sc":
							{
								Console.WriteLine("生成服务器代码");
								break;
							}
						case "-sd":
							{
								Console.WriteLine("生成服务器数据");
								break;
							}
						default:
							{
								showHelp();
								break;
							}
					}
				}

				Current.Shutdown(0);
			}
		}

		private void showHelp()
		{
			Console.WriteLine("┌────────────┐");
			Console.WriteLine("│ -h     帮助            │");
			Console.WriteLine("│ -cc    生成客户端代码  │");
			Console.WriteLine("│ -cd    生成客户端数据  │");
			Console.WriteLine("│ -sc    生成服务器代码  │");
			Console.WriteLine("│ -sd    生成服务器数据  │");
			Console.WriteLine("└────────────┘");
		}
	}
}
