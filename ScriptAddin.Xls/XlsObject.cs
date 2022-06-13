using Excel = Microsoft.Office.Interop.Excel;


namespace ScriptAddin
{
	public class XlsObject
	{

		public XlsObject(Excel.Application app) {
			App = app;
			Book = app.ActiveWorkbook;
			Sheet = (Excel.Worksheet)app.ActiveSheet;
			Cells = Sheet.Cells;
			Sel = (Excel.Range)app.Selection;
		}

		public Excel.Application App { get; private set; }
		public Excel.Workbook Book { get; private set; }
		public Excel.Worksheet Sheet { get; private set; }
		public Excel.Range Cells { get; private set; }
		public Excel.Range Sel { get; private set; }
	}
}
