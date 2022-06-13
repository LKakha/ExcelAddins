using Microsoft.ClearScript;
using System;
using System.Windows.Forms;

namespace ScriptAddin.Engines
{
	internal class EngineBase<T> : IEngine where T : Microsoft.ClearScript.Windows.WindowsScriptEngine, new()
	{
		public ScriptType Type { get; set; }
		public string SyntaxHighlightingName { get; set; }
		const HostItemFlags flags = HostItemFlags.DirectAccess;

		public void Execute(string code, XlsObject xls) {
			try {
				using (var engine = new T()) {
					engine.AddHostObject("App", flags, xls.App);
					engine.AddHostObject("Book", flags, xls.App.ActiveWorkbook);
					engine.AddHostObject("Sheet", flags, xls.Sheet);
					engine.AddHostObject("Cells", flags, xls.Sheet.Cells);
					engine.AddHostObject("Sel", flags, xls.Sel);

					engine.AddHostObject("Ext", new ScriptExtension());
					engine.AddHostObject("host", new HostFunctions());
					engine.AddHostObject("clr", new HostTypeCollection("mscorlib", "System", "System.Core"));

					engine.Execute(code);
					engine.CollectGarbage(false);
				}
			}
			catch (ScriptEngineException ex) {
				throw new Exception(ex.ErrorDetails);
			}
			catch (Exception ex) {
				MessageBox.Show(ex.Message);
			}
		}
	}
}
