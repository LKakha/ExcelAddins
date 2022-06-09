using Excel = Microsoft.Office.Interop.Excel;


namespace ScriptAddin
{
	public class XlsObject
	{

		public XlsObject(Excel.Application app) {
			App = app;
			Book = app.ActiveWorkbook;
			Sheet = app.ActiveSheet as Excel.Worksheet;
			Cells = Sheet.Cells;
			Sel = app.Selection as Excel.Range;
		}

		public Excel.Application App { get; private set; }
		public Excel.Workbook Book { get; private set; }
		public Excel.Worksheet Sheet { get; private set; }
		public Excel.Range Cells { get; private set; }
		public Excel.Range Sel { get; private set; }
	}
}
