using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool.Generaters
{
	class GenClientBolo
	{
		public static string _tab = "    ";

		public static void generater(ISheet sheet, List<ColData> members)
		{
			int rowCount = GenClientData.convertCellType(sheet);

			string cscriptPath = Properties.Settings.Default.cscriptPath;
			StreamWriter f = new StreamWriter(System.IO.Path.Combine(cscriptPath, sheet.SheetName + "Config.bolos"), false);
			f.Write(string.Format(""
				+ "var {0}Table = {{\n"
				+ "    \"ConfigDefine\" : {{\n"
				+ "{1}"
				+ "    }}<{0}Config>,\n"
				+ "    //--datas begin--\n"
				+ "    \"-1\" : {{\n"
				+ "{2}"
				+ "    }}\n"
				+ "{3}"
				+ "}};\n"
				+ "\n"
				+ "int {0}Table_size() {{\n"
				+ "    return {4};\n"
				+ "}}\n"
				+ "\n"
				+ "class {0}Config {0}Table_get(String key) {{\n"
				+ "    if (key.type() == \"int\") {{\n"
				+ "        key = key.toString();\n"
				+ "    }}\n"
				+ "    if (!{0}Table.contains(key)) {{\n"
				+ "        key = \"-1\";\n"
				+ "    }}\n"
				+ "    return {0}Table[key];\n"
				+ "}}\n"
				, sheet.SheetName
				, genDefineConfig(members)
				, genDefaultConfig(members)
				, genDatasList(sheet, members, rowCount)
				, rowCount - Generater.headerRowCount
				));
			f.Close();
		}

		private static string genDefineConfig(List<ColData> members)
		{
			string ret = "";
			foreach (ColData d in members)
			{
				ret += _tab + _tab + d.vname + " : ";
				if (d.isCombineArr || d.isCommaArr)
				{
					ret += "{";
					ret += GenClientCpp.defaultValue(d.vtype);
					ret += "}";
				}
				else
				{
					ret += GenClientCpp.defaultValue(d.vtype) + "";
				}

				if (members.IndexOf(d) == members.Count - 1)
				{
					ret += "\n";
				}
				else
				{
					ret += ",\n";
				}
			}
			return ret;
		}

		private static string genDefaultConfig(List<ColData> members)
		{
			string ret = "";
			foreach (ColData d in members)
			{
				ret += _tab + _tab + d.vname + " : ";
				if (d.isCombineArr || d.isCommaArr)
				{
					ret += "{}";
				}
				else
				{
					ret += GenClientCpp.defaultValue(d.vtype) + "";
				}

				if (members.IndexOf(d) == members.Count - 1)
				{
					ret += "\n";
				}
				else
				{
					ret += ",\n";
				}
			}
			return ret;
		}

		private static string genDatasList(ISheet sheet, List<ColData> members, int rowCount)
		{
			string ret = "";
			ColData cd = null;
			for (int i = Generater.headerRowCount; i < rowCount; i++)
			{
				IRow r = sheet.GetRow(i);
				for (int j = 0; j < members.Count; j++)
				{
					cd = members[j];
					string v = r.GetCell(cd.column).StringCellValue;
					if (j == 0)
					{
						ret += _tab + string.Format("\"{0}\" : {{\n", v);
					}
					ret += _tab + _tab + cd.vname + " : ";

					if (cd.isCommaArr)
					{
						string[] arr = v.Split(',');
						ret += "{\n";
						int idx = 0;
						foreach (string e in arr)
						{
							ret += _tab + _tab + _tab;
							if (cd.vtype == "str")
							{
								ret += "\"" + e + "\"";
							}
							else
							{
								ret += e;
							}

							if (idx == arr.Count()-1)
							{
								ret += "\n";
							}
							else
							{
								ret += ",\n";
							}
							idx++;
						}
						ret += _tab + _tab + "}";
					}
					else if (cd.isCombineArr)
					{
						ret += "{\n";
						for (int k = cd.column; k < cd.column + cd.arrLen; k++)
						{
							v = r.GetCell(k).StringCellValue;
							ret += _tab + _tab + _tab;
							if (cd.vtype == "str")
							{
								ret += "\"" + v + "\"";
							}
							else
							{
								ret += v;
							}

							if (k == cd.column + cd.arrLen - 1)
							{
								ret += "\n";
							}
							else
							{
								ret += ",\n";
							}
						}
						ret += _tab + _tab + "}";
					}
					else
					{
						if (cd.vtype == "str")
						{
							ret += "\"" + v + "\"";
						}
						else
						{
							ret += v;
						}
					}

					if (j == members.Count - 1)
					{
						ret += "\n";
					}
					else
					{
						ret += ",\n";
					}
				}

				ret += _tab + "}";
				if (i == rowCount - 1)
				{
					ret += "\n";
				}
				else
				{
					ret += ",\n";
				}	
			}
			return ret;
		}
	}
}
