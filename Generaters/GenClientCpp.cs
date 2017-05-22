using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool.Generaters
{
	class GenClientCpp
	{
		public static string _tab = "    ";

		public static void generater(string sheetName, List<ColData> members)
		{
			genHeader(sheetName, members);
			genCpp(sheetName, members);
		}

		private static void genHeader(string sheetName, List<ColData> members)
		{
			string ccodePath = Properties.Settings.Default.ccodePath;
			StreamWriter f = new StreamWriter(System.IO.Path.Combine(ccodePath, sheetName + "Config.h"), false);
			f.Write(string.Format(""
				+ "#ifndef __{0}CONFIG_H__\n"
				+ "#define __{0}CONFIG_H__\n"
				+ "\n"
				+ "#include <gstl.h>\n"
				+ "\n"
				+ "class {1}Config {{\n"
				+ "public:\n"
				+ "{2}\n"
				+ "    {1}Config();\n"
				+ "}};\n"
				+ "\n"
				+ "class {1}Table {{\n"
				+ "public:\n"
				+ "    static {1}* getInstance();\n"
				+ "\n"
				+ "    static {1}Config* get({3} key);\n"
				+ "    static void load();\n"
				+ "    static void clear();\n"
				+ "    int getSize() const;\n"
				+ "\n"
				+ "private:\n"
				+ "    static {1}Table* instance;\n"
				+ "    HashMap<{3}, {1}Config> datas;\n"
				+ "    {1}Config* defaultConfig;\n"
				+ "    int size;\n"
				+ "}};\n"
				+ "\n"
				+ "#endif /*__{0}CONFIG_H__*/"
				, sheetName.ToUpper()
				, sheetName
				, genMemberDeclareStr(members)
				, typeStr(members[0].vtype)
				));
			f.Close();
		}

		private static void genCpp(string sheetName, List<ColData> members)
		{
			string ccodePath = Properties.Settings.Default.ccodePath;
			StreamWriter f = new StreamWriter(System.IO.Path.Combine(ccodePath, sheetName + "Config.cpp"), false);
			f.Write(string.Format(""
				+ "#include \"{0}Config.h\"\n"
				+ "#include \"ResLoader.h\"\n"
				+ "\n"
				+ "{0}Config::{0}Config()\n"
				+ "{1}"
				+ "{{}}\n"
				+ "\n"
				+ "{0}Config* {0}Table::get({4} key) {{\n"
				+ "    if (size == 0) {{\n"
				+ "        load();\n"
				+ "    }}\n"
				+ "    HashMap<{4}, {0}Config>::iterator it = datas.find(key);\n"
				+ "    if (it != datas.end()) {{\n"
				+ "    return it->second;\n"
				+ "    }}\n"
				+ "    return defaultConfig;\n"
				+ "}}\n"
				+ "\n"
				+ "void {0}Table::load() {{\n"
				+ "    s32 len;\n"
				+ "    c8 * file = ResLoader::loadFile(\"res/config/{0}Config.bin\", len);\n"
				+ "    iobuf buf(file, len);\n"
				+ "    size = buf.readInt32();\n"
				+ "    for (int i = 0; i < size; i++) {{\n"
				+ "        {0}Config conf;\n"
				+ "{2}"
				+ "        datas[conf.{3}] = conf;\n"
				+ "    }}\n"
				+ "}}\n"
				+ "\n"
				+ "void {0}Table::clear() {{\n"
				+ "    datas.clear();\n"
				+ "}}\n"
				, sheetName
				, genConstructStr(members)
				, genReadMemberStr(members)
				, members[0].vname
				, typeStr(members[0].vtype)
				));
			f.Close();
		}

		public static string genMemberDeclareStr(List<ColData> members)
		{
			string ret = "";
			foreach (ColData d in members)
			{
				bool isArr = d.isCommaArr || d.isCombineArr;
				string vtypeStr = (isArr ? d.vtype + "[]" : d.vtype);
				ret += _tab + typeStr(vtypeStr) + " " + d.vname + ";\n";
			}
			return ret;
		}

		public static string genConstructStr(List<ColData> members)
		{
			string ret = "";
			foreach (ColData d in members)
			{
				bool isArr = d.isCommaArr || d.isCombineArr;
				if (!isArr)
				{
					if (d.column == 0 && d.vtype == "int")
					{
						ret += (ret == "" ? ":" : ",");
						ret += d.vname + "(-1)\n";
					}
					else
					{
						ret += (ret == "" ? ":" : ",");
						ret += d.vname + "(" + defaultValue(d.vtype) + ")\n";
					}
				}
			}
			return ret;
		}

		public static string genReadMemberStr(List<ColData> members)
		{
			string ret = "";
			foreach (ColData d in members)
			{
				bool isArr = d.isCommaArr || d.isCombineArr;
				ret += readMemberStr(_tab + _tab, "conf." + d.vname, d.vtype, isArr);
			}
			return ret;
		}

		public static string typeStr(string t)
		{
			switch (t)
			{
				case "int":
					return "s32";
				case "int[]":
					return "ArrayList<s32>";
				case "str":
					return "wstring";
				case "str[]":
					return "ArrayList<wstring>";
				case "float":
					return "float";
				case "float[]":
					return "ArrayList<f32>";
				default:
					return t;
			}
		}

		private static string readMemberStr(string tab, string vname, string vtype, bool isArr)
		{
			string ret = "";
			if (isArr)
			{
				ret = tab + "len = buf.readInt32();\n"
					+ tab + vname + ".resize(len);\n"
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
					ret += "readUTF();\n";
				}
				else
				{
					ret += "[type error: " + vtype + "];\n";
				}
			}
			return ret;
		}

		public static string defaultValue(string t)
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
					return "0";
			}
		}

	}
}
