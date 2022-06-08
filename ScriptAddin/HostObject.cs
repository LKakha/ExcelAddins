using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace ScriptAddin
{
	public class HostObject
	{

		public HostObject(Excel.Application app) {
			App = app;
			Book = app.ActiveWorkbook;
			Sheet = app.ActiveSheet as Excel.Worksheet;
			Sel = app.Selection as Excel.Range;
		}

		public Excel.Application App { get; private set; }
		public Excel.Workbook Book { get; private set; }
		public Excel.Worksheet Sheet { get; private set; }
		public Excel.Range Sel { get; private set; }
	}
}
