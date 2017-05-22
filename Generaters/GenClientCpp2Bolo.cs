using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool.Generaters
{
	class GenClientCpp2Bolo
	{
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
				+ "#include \"BoloObject.h\"\n"
				+ "\n"
				+ "using namespace bs;\n"
				+ "\n"
				+ "class {1}Config : public BoloObject {{\n"
				+ "    using Self = {1}Config;\n"
				+ "public:\n"
				+ "{2}\n\n"
				+ "{4}\n\n"
				+ "    {1}Config();\n"
				+ "    virtual const string& getClassName() const {{\n"
				+ "        static string name = \"{1}Config\";\n"
				+ "        return name;\n"
				+ "    }}\n"
				+ "    static void registerReflection(s32 id);\n"
				+ "}};\n"
				+ "\n"
				+ "class {1}Table : public BoloObject {{\n"
				+ "    using Self = {1}Table;\n"
				+ "public:\n"
				+ "    static {1}Table* getInstance();\n"
				+ "\n"
				+ "    {1}Config* get({3} key);\n"
				+ "    void load();\n"
				+ "    void clear();\n"
				+ "    int getSize() const;\n"
				+ "\n"
				+ "    BoloVar bolo_get(bolo_stack stack);\n"
				+ "    virtual const string& getClassName() const {{\n"
				+ "        static string name = \"{1}Table\";\n"
				+ "        return name;\n"
				+ "    }}\n"
				+ "    static void registerReflection(s32 id);\n"
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
				, GenClientCpp.genMemberDeclareStr(members)
				, GenClientCpp.typeStr(members[0].vtype)
				, genGetterFuncDeclareStr(members)
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
				+ "{5}"
				+ "\n"
				+ "void {0}Config::registerReflection(s32 id) {{\n"
				+ "{6}"
				+ "}}\n"
				+ "\n\n"
				+ "{0}Table* {0}Table::instance = nullptr;\n"
				+ "{0}Table* {0}Table::getInstance() {{\n"
				+ "    if (!instance) {{\n"
				+ "        BoloObject::initScriptLib<{0}Config>();\n"
				+ "        BoloObject::initScriptLib<{0}Table>();\n"
				+ "        instance = new {0}Table();\n"
				+ "        BoloVM::registerEnterClass(instance, instance->getClassName());\n"
				+ "    }}\n"
				+ "    return instance;\n"
				+ "}}\n"
				+ "\n"
				+ "const {0}Config& {0}Table::get({4} key) {{\n"
				+ "    if (size == 0) {{\n"
				+ "        load();\n"
				+ "    }}\n"
				+ "    HashMap<{4}, {0}Config>::iterator it = datas.find(key);\n"
				+ "    if (it != datas.end()) {{\n"
				+ "        return it->second;\n"
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
				+ "\n"
				+ "    defaultConfig = new MedicineConfig();\n"
				+ "}}\n"
				+ "\n"
				+ "void {0}Table::clear() {{\n"
				+ "    for (auto it : datas) {{\n"
				+ "        delete it.second;\n"
				+ "    }}\n"
				+ "    datas.clear();\n"
				+ "}}\n"
				+ "\n"
				+ "int {0}Table::getSize() const {{\n"
				+ "    if (size == 0) {{\n"
				+ "        (({0}Table*)this)->load();\n"
				+ "    }}\n"
				+ "    return size;\n"
				+ "}}\n"
				+ "\n"
				+ "BoloVar {0}Table::bolo_get(bolo_stack) {{\n"
				+ "    int key = bolo_int(stack);\n"
				+ "    return bolo_create(stack, get(key), false);\n"
				+ "}}\n"
				+ "\n"
				+ "void {0}Table::registerReflection(s32 id) {{\n"
				+ "    BoloRegisterFunction(\"get\", bolo_get);\n"
				+ "    BoloRegisterPropReadOnly(\"size\", getSize();\n"
				+ "}}"
				, sheetName
				, GenClientCpp.genConstructStr(members)
				, GenClientCpp.genReadMemberStr(members)
				, members[0].vname
				, GenClientCpp.typeStr(members[0].vtype)
				, genGetterFuncSynthesize(members, sheetName)
				, genRegisterReflection(members)
				));
			f.Close();
		}

		public static string genGetterFuncDeclareStr(List<ColData> members)
		{
			string ret = "";
			foreach (ColData d in members)
			{
				if (d.vtype == "str")
				{
					ret += GenClientCpp._tab + "const wstring& get" + toTitle(d.vname) + "() const { return " + d.vname + ";}\n";
				}
				else if (d.isCombineArr || d.isCommaArr)
				{
					ret += GenClientCpp._tab + "BoloVar get" + toTitle(d.vname) + "Size(bolo_statck stack);\n";
					ret += GenClientCpp._tab + "BOloVar get" + toTitle(d.vname) + "At(bolo_stack stack);\n";
				}
				else
				{
					ret += GenClientCpp._tab + GenClientCpp.typeStr(d.vtype) + " get" + toTitle(d.vname) + "() const { return " + d.vname + "; }\n";
				}
			}
			return ret;
		}

		public static string genGetterFuncSynthesize(List<ColData> members, string classname)
		{
			string ret = "";
			foreach (ColData d in members)
			{
				if (d.isCombineArr || d.isCommaArr)
				{
					ret += string.Format(""
						+ "BoloVar {0}Config::bolo_{1}Size(bolo_stack stack) const {{\n"
						+ "    return {2}.size();\n"
						+ "}}\n"
						+ "\n"
						+ "BoloVar {0}Config::bolo_{1}At(bolo_stack stack) const {{\n"
						+ "    int idx = bolo_int(stack);\n"
						+ "    if (idx >= 0 && idx < {2}.size()) {{\n"
						+ "        return {2}[idx];\n"
						+ "    }}\n"
						+ "    return {3};\n"
						+ "}}\n"
						, classname, toTitle(d.vname), d.vname, defaultErrorVal(d.vtype));
				}
			}
			return ret;
		}

		public static string genRegisterReflection(List<ColData> members)
		{
			string ret = "";
			foreach (ColData d in members)
			{
				if (d.isCombineArr || d.isCommaArr)
				{
					ret += GenClientCpp._tab + "BoloRegisterFunction(\"" + d.vname + "_size\", bolo_" + toTitle(d.vname) + "Size);\n";
					ret += GenClientCpp._tab + "BoloRegisterFunction(\"" + d.vname + "_at\", bolo_" + toTitle(d.vname) + "At);\n";
				}
				else
				{
					ret += GenClientCpp._tab + "BoloRegisterPropReadOnly(\"" + d.vname + "\", get" + toTitle(d.vname) + ");\n";
				}

			}
			return ret;
		}

		public static string toTitle(string s)
		{
			return System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(s);
		}

		private static string defaultErrorVal(string t)
		{
			switch (t)
			{
				case "int":
					return "-999";
				case "str":
					return "\"error\"";
				case "float":
					return "-999";
				default:
					return "-999";
			}
		}

	}
}
