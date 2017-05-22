using MiscUtil.Conversion;
using MiscUtil.IO;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using ExcelTool.Generaters;

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
		public static int headerRowCount = 4;

        public static bool genClientCode(ISheet sheet)
        {
			List<ColData> members = parseStruct(sheet, "C");
			if (members.Count == 0)
			{
				return false;
			}

			switch (MainWindow.getSheetOutType(sheet.SheetName))
			{
				case COutType.CPP:
					{
						GenClientCpp.generater(sheet.SheetName, members);
						break;
					}
				case COutType.BOLO:
					{
						GenClientBolo.generater(sheet, members);
						break;
					}
				case COutType.CPP2BOLO:
					{
						GenClientCpp2Bolo.generater(sheet.SheetName, members);
						break;
					}
				default:
					break;
			}

			return true;
		}

        public static bool genClientData(ISheet sheet)
        {
			List<ColData> members = parseStruct(sheet, "C");
			if (members.Count == 0)
			{
				return false;
			}

			GenClientData.generater(sheet, members);
			
			return true;
		}

		
		public static bool genServerCode(ISheet sheet)
        {
			return true;
        }


        public static bool genServerData(ISheet sheet)
        {
			return true;
		}


		

		private static List<ColData> parseStruct(ISheet sheet, string cs)
		{
			IRow r2 = sheet.GetRow(1);  //字段名
			IRow r3 = sheet.GetRow(2);  //字段类型
			IRow r4 = sheet.GetRow(3);  //前端后端

			List<ColData> members = new List<ColData>();
			ColData cd = null;
			string savedName = "";

			for (int i = 0; i < r2.LastCellNum; i++)
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
		
    }
}
