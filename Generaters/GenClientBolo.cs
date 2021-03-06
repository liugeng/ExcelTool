﻿using NPOI.SS.UserModel;
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
		const int WRITE_WITH_FORMAT = 0;
		public static string _tab = "    ";

		public static void generater(ISheet sheet, List<ColData> members)
		{
			int rowCount = GenClientData.convertCellType(sheet);

			string cscriptPath = Properties.Settings.Default.cscriptPath;
			StreamWriter f = new StreamWriter(System.IO.Path.Combine(cscriptPath, sheet.SheetName + "Config.bolos"), false);
			f.Write(string.Format(""
				+ "{4}\n"
				+ "var {0}Table = {{\n"
				+ "{1}<{0}Config>,\n"
				+ "{2}"
				+ "}};\n"
				+ "\n"
				+ "int {0}Table_size() {{\n"
				+ "    return {3};\n"
				+ "}}\n"
				+ "\n"
				+ "class {0}Config {0}Table_get(String key) {{\n"
				+ "    if (key.type() == \"int\") {{\n"
				+ "        key = key.toString();\n"
				+ "    }}\n"
				+ "    if (!{0}Table.keys().contains(key)) {{\n"
				+ "        print(\"Key not found in {0}Table: \" + key);\n"
				+ "        key = \"-1\";\n"
				+ "    }}\n"
				+ "    return {0}Table[key];\n"
				+ "}}\n"
				, sheet.SheetName
				//, genDefineConfig(members)
				, genDefaultConfig(members)
				, genDatasList(sheet, members, rowCount)
				, rowCount - Generater.headerRowCount
				, Generater.fileHeader
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
#if WRITE_WITH_FORMAT //格式化排版
			string ret = _tab + "\"-1\" : {\n";
			foreach (ColData d in members)
			{

				ret += _tab + _tab + d.vname + " : ";
				if (d.isCombineArr || d.isCommaArr)
				{
					ret += "{}";
				}
				else if (members.IndexOf(d) == 0 && d.vtype == "int")
				{
					ret += "-1";
				}
				else
				{
					ret += GenClientCpp.defaultValue(d.vtype);
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
			ret += _tab + "}"
#else
			string ret = _tab + "\"-1\" : { ";
			foreach (ColData d in members)
			{
				ret += d.vname + ":";
				if (d.isCombineArr || d.isCommaArr)
				{
					ret += "{}";
				}
				else if (members.IndexOf(d) == 0 && d.vtype == "int")
				{
					ret += "-1";
				}
				else
				{
					ret += GenClientCpp.defaultValue(d.vtype);
				}

				if (members.IndexOf(d) < members.Count - 1)
				{
					ret += ", ";
				}
			}
			ret += "}";
#endif

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

#if WRITE_WITH_FORMAT //格式化排版
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

#else //内容排版紧缩
					if (j == 0)
					{
						ret += _tab + string.Format("\"{0}\" : {{", v);
					}
					ret += cd.vname + ":";

					if (cd.isCommaArr)
					{
						string[] arr = v.Split(',');
						ret += "{";
						int idx = 0;
						foreach (string e in arr)
						{
							if (cd.vtype == "str")
							{
								ret += "\"" + e + "\"";
							}
							else
							{
								ret += e;
							}

							if (idx < arr.Count() - 1)
							{
								ret += ",";
							}
							idx++;
						}
						ret += "}";
					}
					else if (cd.isCombineArr)
					{
						ret += "{";
						for (int k = cd.column; k < cd.column + cd.arrLen; k++)
						{
							v = r.GetCell(k).StringCellValue;
							if (cd.vtype == "str")
							{
								ret += "\"" + v + "\"";
							}
							else
							{
								ret += v;
							}

							if (k < cd.column + cd.arrLen - 1)
							{
								ret += ",";
							}
						}
						ret += "}";
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

					if (j < members.Count - 1)
					{
						ret += ", ";
					}
				}

				ret += "}";
#endif

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
