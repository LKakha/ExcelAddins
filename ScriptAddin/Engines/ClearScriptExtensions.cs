using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using System.Runtime.InteropServices.WindowsRuntime;

namespace ScriptAddin.Engines
{
	public class ScriptExtension : Microsoft.ClearScript.ExtendedHostFunctions
	{
		public DialogResult MsgBox(object Msg, MessageBoxButtons Button = MessageBoxButtons.OK) {
			return MessageBox.Show(Msg?.ToString(), "ScriptAddin", Button);
		}
		public DialogResult alert(object Msg, MessageBoxButtons Button = MessageBoxButtons.OK) {
			return MsgBox(Msg, Button);
		}
		public ScriptItem[] SelectedCells() {
			//var excel= (Excel.Application)ExcelDnaUtil.Application;
			//var selection = ((Excel.Range)excel.Selection);
			var cells = new List<ScriptItem>();
			//var count = selection.Count;
			//for (var i=1; i<=count; i++) {
			//	cells.Add((Excel.Range)selection.Item[i]);
			//}
			//return cells.ToArray();
			//var sheet = (Excel.Worksheet)excel.ActiveSheet;
			for (var i=1; i<=10000; i++) {
				cells.Add(new ScriptItem());
			}
			return cells.ToArray();
		}
	}
}
