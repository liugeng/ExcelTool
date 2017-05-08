using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace ExcelTool
{
	class ColData
	{
		public string vname;
		public string vtype;
		public bool isCommaArr;		//用逗号分隔的数组
		public bool isCombineArr;	//多列组成的数组
		public int arrLen;
		public int column;
	}

    class Generater
    {
		static int headerRowCount = 4;

        public static void genClientCode(ISheet sheet)
        {
            IRow r1 = sheet.GetRow(0);  //备注
            IRow r2 = sheet.GetRow(1);  //字段名
            IRow r3 = sheet.GetRow(2);  //字段类型
            IRow r4 = sheet.GetRow(3);  //前端后端

			List<ColData> members = parseStruct(sheet, "C");
			string saveName = "";
            string memStr = "";
            string constructStr = "";
			string readMemStr = "";

			//todo: 使用members生成memStr和constructStr
			foreach(ColData d in members)
			{
				readMemStr += readMemberStr("\t\t", "conf." + d.vname, d.vtype, d.isCommaArr || d.isCombineArr);
			}


            for (int i = r2.FirstCellNum; i < r2.LastCellNum; i++)
            {
				if (!r4.GetCell(i).StringCellValue.Contains("C"))
				{
					continue;
				}
                string name = r2.GetCell(i).StringCellValue;
                string vtype = r3.GetCell(i).StringCellValue;
                if (name.Contains("[")) //组合数组
                {
                    if (saveName == name)
                    {
                        continue;
                    }
                    memStr += "\t" + typeStr(vtype + "[]") + " " + name.Substring(0, name.IndexOf("[")) + ";\n";
					//constructStr += (constructStr == "" ? ":" : ",");
					//constructStr += name.Substring(0, name.IndexOf("[")) + "(" + defaultValue(vtype + "[]") + ")\n";
				}
				else
                {
                    memStr += "\t" + typeStr(vtype) + " " + name + ";\n";
					if (i == r2.FirstCellNum && vtype == "int")
					{
						constructStr += (constructStr == "" ? ":" : ",");
						constructStr += name + "(-1)\n";
					}
					else if (!vtype.Contains("["))
					{
						constructStr += (constructStr == "" ? ":" : ",");
						constructStr += name + "(" + defaultValue(vtype) + ")\n";
					}
                }
                saveName = name;
            }


            string ccodePath = Properties.Settings.Default.ccodePath;
            StreamWriter f = new StreamWriter(System.IO.Path.Combine(ccodePath, sheet.SheetName + "Config.h"), false);
            f.Write(string.Format(""
                + "#ifndef __{0}CONFIG_H__\n"
                + "#define __{0}CONFIG_H__\n"
                + "\n"
                + "class {1}Config {{\n"
                + "public:\n"
                + "{2}\n"
                + "	{1}Config();\n"
                + "}};\n"
                + "\n"
                + "class {1}Table {{\n"
                + "public:\n"
                + "	static const {1}Config& get(s32 key);\n"
                + "	static void load();\n"
                + "	static void clear();\n"
				+ "	static int size;\n"
				+ "private:\n"
                + "	static HashMap<s32, {1}Config> datas;\n"
                + "	static {1}Config defaultConfig;\n"
                + "}};\n"
                + "\n"
                + "#endif /*__{0}CONFIG_H__*/", sheet.SheetName.ToUpper(), sheet.SheetName, memStr));
            f.Close();


			f = new StreamWriter(System.IO.Path.Combine(ccodePath, sheet.SheetName + "Config.cpp"), false);
			f.Write(string.Format(""
				+ "#include \"{0}Config.h\"\n"
				+ "\n"
				+ "{0}Config::{0}Config()\n"
				+ "{1}"
				+ "{{}}\n"
				+ "\n"
				+ "HashMap<s32, {0}Config> {0}Table::datas;\n"
				+ "{0}Config {0}Table::defaultConfig;\n"
				+ "int {0}Table::size = 0;\n"
				+ "\n"
				+ "const {0}Config& {0}Table::get(s32 key) {{\n"
				+ "	if (size == 0) {{\n"
				+ "		load();\n"
				+ "	}}\n"
				+ "	HashMap<s32, {0}Config>::iterator it = datas.find(key);\n"
				+ "	if (it != datas.end()) {{\n"
				+ "		return it->second;\n"
				+ "	}}\n"
				+ "	return defaultConfig;\n"
				+ "}}\n"
				+ "\n"
				+ "void {0}Table::load() {{\n"
				+ "	s32 len;\n"
				+ "	c8 * file = ResLoader::loadFile(\"{0}Config.bin\", len);\n"
				+ "	iobuf buf(file, len);\n"
				+ "	size = buf.readInt32();\n"
				+ "	for (int i = 0; i < size; i++) {{\n"
				+ "		{0}Config conf;\n"
				+ "{2}"
				+ "	}}\n"
				+ "}}\n"
				+ "\n"
				+ "void {0}Table::clear() {{\n"
				+ "	datas.clear();\n"
				+ "}}\n"
				+ "", sheet.SheetName, constructStr, readMemStr));
			f.Close();
		}

        public static void genClientData(ISheet sheet)
        {
			int rowCount = convertCellType(sheet);

			IRow r2 = sheet.GetRow(1);  //字段名
			IRow r3 = sheet.GetRow(2);  //字段类型
			IRow r4 = sheet.GetRow(3);  //前端后端

			List<ColData> members = parseStruct(sheet, "C");

			string filename = Path.Combine(Properties.Settings.Default.cdataPath, sheet.SheetName + "Config.bin");
			BinaryWriter bw = new BinaryWriter(new FileStream(filename, FileMode.Create));
			ColData cd = null;

			bw.Write(rowCount- headerRowCount);
			for (int i = headerRowCount; i < rowCount; i++)
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
						for (int k = cd.column; k < cd.column+cd.arrLen; k++)
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

		private static string readMemberStr(string tab, string vname, string vtype, bool isArr)
		{
			string ret = "";
			if (isArr)
			{
				ret = tab + "len = buf.readInt32();\n"
					+ tab + "for (int j = 0; j < len; j++) {\n"
					+ readMemberStr(tab + "\t", vname + "[j]", vtype, false)
					+ tab + "}\n";
			}
			else
			{
				ret = tab + vname + " = buf.";
				if (vtype == "int")
				{
					ret += "readInt32();\n";
				}
				else if (vtype == "float")
				{
					ret += "readFloat();\n";
				}
				else if (vtype == "str")
				{
					ret += "readString();\n";
				}
				else
				{
					ret += "[type error: " + vtype + "];\n";
				}
			}
			return ret;
		}

		public static void genServerCode(ISheet sheet)
        {

        }


        public static void genServerData(ISheet sheet)
        {

        }



        private static string typeStr(string t)
        {
            switch (t)
            {
                case "int":
                    return "s32";
                case "int[]":
                    return "ArrayList<s32>";
                case "str":
                    return "string";
                case "str[]":
                    return "ArrayList<string>";
                case "float":
                    return "float";
                case "float[]":
                    return "ArrayList<f32>";
                default:
                    return t;
            }
        }

        private static string defaultValue(string t)
        {
            switch (t)
            {
                case "int":
                    return "0";
                case "str":
                    return "\"\"";
                case "float":
                    return "0.0";
                default:
                    return t;
            }
        }

		private static List<ColData> parseStruct(ISheet sheet, string cs)
		{
			IRow r2 = sheet.GetRow(1);  //字段名
			IRow r3 = sheet.GetRow(2);  //字段类型
			IRow r4 = sheet.GetRow(3);  //前端后端

			List<ColData> members = new List<ColData>();
			ColData cd = null;
			string savedName = "";

			for (int i = r2.FirstCellNum; i < r2.LastCellNum; i++)
			{
				if (!r4.GetCell(i).StringCellValue.Contains(cs))
				{
					continue;
				}
				string vname = r2.GetCell(i).StringCellValue;
				string vtype = r3.GetCell(i).StringCellValue;

				if (vname.Contains("[")) //组合数组
				{
					if (savedName != vname)
					{
						cd = new ColData()
						{
							vname = vname.Substring(0,vname.IndexOf("[")),
							vtype = vtype,
							isCommaArr = false,
							isCombineArr = true,
							arrLen = 1,
							column = i
						};
						members.Add(cd);
					}
					else
					{
						cd.arrLen++;
					}
				}
				else
				{
					cd = new ColData()
					{
						vname = vname,
						vtype = vtype,
						isCommaArr = vtype.Contains("["),
						isCombineArr = false,
						arrLen = 0,
						column = i
					};
					if (vtype.Contains("["))
					{
						cd.vtype = cd.vtype.Substring(0, vtype.IndexOf("["));
					}
					members.Add(cd);
				}

				savedName = vname;
			}

			return members;
		}

		private static void writeMember(BinaryWriter bw, string vtype, string v, string errmsg)
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
				bw.Write(v);
			}
		}

		private static int convertCellType(ISheet sheet)
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
						if (i > headerRowCount)
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
    }
}
