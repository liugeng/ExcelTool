using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool
{
    class Generater
    {
        public static void genClientCode(ISheet sheet)
        {
            IRow r1 = sheet.GetRow(0);  //备注
            IRow r2 = sheet.GetRow(1);  //字段名
            IRow r3 = sheet.GetRow(2);  //字段类型
            IRow r4 = sheet.GetRow(3);  //前段后端

            string saveName = "";
            string initStr = "";
            string memStr = "";

            for (int i = r2.FirstCellNum; i < r2.LastCellNum; i++)
            {
                string name = r2.GetCell(i).StringCellValue;
                string vtype = r3.GetCell(i).StringCellValue;
                if (name.Contains("[]"))
                {
                    if (saveName == name)
                    {
                        continue;
                    }
                    initStr += (initStr == "" ? "\t:" : "\t,");
                    memStr += "\t" + typeStr(vtype + "[]") + " " + name.Substring(0, name.IndexOf("[")) + ";\n";
                    initStr += name.Substring(0, name.IndexOf("[")) + "(" + defaultValue(vtype + "[]") + ")\n";
                }
                else
                {
                    initStr += (initStr == "" ? "\t:" : "\t,");
                    memStr += "\t" + typeStr(vtype) + " " + name + ";\n";
                    if (i == r2.FirstCellNum && vtype == "int")
                    {
                        initStr += name + "(-1)\n";
                    }
                    else
                    {
                        initStr += name + "(" + defaultValue(vtype) + ")\n";
                    }
                }
                saveName = name;
            }


            string outputPath = Properties.Settings.Default.outputPath;
            StreamWriter f = new StreamWriter(System.IO.Path.Combine(outputPath, sheet.SheetName + "Config.hpp"), false);
            f.Write(string.Format(""
                + "#ifndef __{0}CONFIG_H__\n"
                + "#define __{0}CONFIG_H__\n"
                + "\n"
                + "class {1}Config {{\n"
                + "public:\n"
                + "{2}\n"
                + "\t{1}Config()\n"
                + "{3}"
                + "\t{{}}\n"
                + "}};\n"
                + "\n"
                + "class {1}Table {{\n"
                + "public:\n"
                + "\tstatic {1}Config* get(s32 key) {{\n"
                + "\t\tif (datas.size() == 0) {{\n"
                + "\t\t\tload();\n"
                + "\t\t\tdefaultConfig = new {1}Config();\n"
                + "\t\t}}\n"
                + "\t\tHashMap<s32, {1}Config*>::iterator it = datas.find(key);\n"
                + "\t\tif (it != datas.end()) {{\n"
                + "\t\t\treturn it->second;\n"
                + "\t\t}}\n"
                + "\t\treturn defaultConfig;\n"
                + "\t}};\n\n"
                + "\tstatic int size() {{ return nSize; }}\n\n"
                + "\tstatic void load() {{\n"
                + "\t}}\n\n"
                + "\tstatic void release() {{\n"
                + "\t}}\n\n"
                + "private:\n"
                + "\tstatic HashMap<s32, {1}Config*> datas;\n"
                + "\tstatic {1}Config* defaultConfig;\n"
                + "\tstatic int nSize;\n"
                + "}};\n"
                + "\n"
                + "HashMap<s32, {1}Config*> {1}Table::datas;\n"
                + "{1}Config* {1}Table::defaultConfig;\n"
                + "int {1}Table::nSize = 0;\n"
                + "\n"
                + "#endif /*__{0}CONFIG_H__*/", sheet.SheetName.ToUpper(), sheet.SheetName, memStr, initStr));
            f.Close();
        }

        public static void genServerCode(ISheet sheet)
        {

        }

        public static void genClientData(ISheet sheet)
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
                case "int[]":
                    return "new ArrayList<s32>";
                case "str":
                    return "\"\"";
                case "str[]":
                    return "new ArrayList<string>";
                case "float":
                    return "float";
                case "float[]":
                    return "new ArrayList<f32>";
                default:
                    return t;
            }
        }
    }
}
