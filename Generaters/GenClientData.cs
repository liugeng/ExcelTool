using MiscUtil.Conversion;
using MiscUtil.IO;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool.Generaters
{
	class GenClientData
	{
		public static void generater(ISheet sheet, List<ColData> members)
		{
			string filename = Path.Combine(Properties.Settings.Default.cdataPath, sheet.SheetName + "Config.bin");

			EndianBinaryWriter bw = new EndianBinaryWriter(EndianBitConverter.Big, new FileStream(filename, FileMode.Create), Encoding.UTF8);

			//BinaryWriter bw = new BinaryWriter(new FileStream(filename, FileMode.Create), System.Text.Encoding.BigEndianUnicode);
			ColData cd = null;

			int rowCount = convertCellType(sheet);

			bw.Write(rowCount - Generater.headerRowCount);
			for (int i = Generater.headerRowCount; i < rowCount; i++)
			{
				IRow r = sheet.GetRow(i);
				for (int j = 0; j < members.Count; j++)
				{
					cd = members[j];
					string v = r.GetCell(cd.column).StringCellValue;
					if (cd.isCommaArr)
					{
						string[] arr = v.Split(',');
						bw.Write((int)arr.Length);
						foreach (string e in arr)
						{
							string errmsg = string.Format("单元格({0},{1})内容错误, 要求类型 {2}: [{3}]", i + 1, Char.ConvertFromUtf32('A' + cd.column), cd.vtype, e);
							writeMember(bw, cd.vtype, e, errmsg);
						}
					}
					else if (cd.isCombineArr)
					{
						bw.Write(cd.arrLen);
						for (int k = cd.column; k < cd.column + cd.arrLen; k++)
						{
							v = r.GetCell(k).StringCellValue;
							string errmsg = string.Format("单元格({0},{1})内容错误, 要求类型 {2}: [{3}]", i + 1, Char.ConvertFromUtf32('A' + cd.column), cd.vtype, v);
							writeMember(bw, cd.vtype, v, errmsg);
						}
					}
					else
					{
						string errmsg = string.Format("单元格({0},{1})内容错误, 要求类型 {2}: [{3}]", i + 1, Char.ConvertFromUtf32('A' + cd.column), cd.vtype, v);
						writeMember(bw, cd.vtype, v, errmsg);
					}

				}
			}

			bw.Close();


			//debug read bin
			//BinaryReader br = new BinaryReader(new FileStream(filename, FileMode.Open));
			//int len = br.ReadInt32();
			//Console.WriteLine(len);
			//for (int i = 0; i < len; i++)
			//{
			//foreach (ColData d in members)
			//{
			//	if (d.isCommaArr || d.isCombineArr)
			//	{
			//		len = br.ReadInt32();
			//		Console.WriteLine(len);
			//		for (int j = 0; j < len; j++)
			//		{
			//			readMember(br, d.vtype);
			//		}
			//	}
			//	else
			//	{
			//		readMember(br, d.vtype);
			//	}
			//}
			//	Console.Write("\n");
			//}
			//br.Close();
		}

		//debug read bin
		private static void readMember(BinaryReader br, string vtype)
		{
			if (vtype == "int")
			{
				int n = br.ReadInt32();
				Console.Write(n);
			}
			else if (vtype == "float")
			{
				float f = br.ReadSingle();
				Console.Write(f);
			}
			else if (vtype == "str")
			{
				string s = br.ReadString();
				Console.Write(s);
			}
			else
			{
				throw new Exception("> 没定义的类型: " + vtype);
			}

			Console.Write(" ");
		}

		public static int convertCellType(ISheet sheet)
		{
			List<string> ids = new List<string>();

			int col = sheet.GetRow(2).LastCellNum;
			for (int i = 0; i < sheet.PhysicalNumberOfRows; i++)
			{
				IRow r = sheet.GetRow(i);
				for (int j = 0; j < col; j++)
				{
					r.GetCell(j).SetCellType(CellType.String);

					if (j == 0)
					{
						string val = r.GetCell(j).StringCellValue;

						//check ids
						if (i > Generater.headerRowCount)
						{
							if (ids.Contains(val))
							{
								throw new Exception("> ID重复: " + val);
							}
							ids.Add(val);
						}

						if (val == "")
						{
							return i;
						}
					}
				}
			}
			return sheet.PhysicalNumberOfRows;
		}

		public static void writeMember(EndianBinaryWriter bw, string vtype, string v, string errmsg)
		{
			if (vtype == "int")
			{
				if (v == "")
				{
					bw.Write(0);
				}
				else
				{
					int n = 0;
					if (int.TryParse(v, out n))
					{
						bw.Write(n);
					}
					else
					{
						//MainWindow.instance.Dispatcher.BeginInvoke(new Action<string, Brush>(MainWindow.instance.addLog), errmsg, new SolidColorBrush(Colors.Red));
						bw.Close();
						throw new Exception("> " + errmsg);
					}
				}
			}
			else if (vtype == "float")
			{
				if (v == "")
				{
					bw.Write(0.0f);
				}
				else
				{
					float n = 0;
					if (float.TryParse(v, out n))
					{
						bw.Write(n);
					}
					else
					{
						//MainWindow.instance.Dispatcher.BeginInvoke(new Action<string, Brush>(MainWindow.instance.addLog), errmsg, new SolidColorBrush(Colors.Red));
						bw.Close();
						throw new Exception("> " + errmsg);
					}
				}
			}
			else if (vtype == "str")
			{
				byte[] bytes = Encoding.UTF8.GetBytes(v);
				bw.Write((short)bytes.Length);
				bw.Write(bytes);
			}
		}
	}
}
