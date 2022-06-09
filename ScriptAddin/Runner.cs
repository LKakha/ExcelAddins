using ScriptAddin.Engines;
using System;
using Excel = Microsoft.Office.Interop.Excel;


namespace ScriptAddin
{
	public class Runner
	{
		readonly TinyIoC.TinyIoCContainer IoC = TinyIoC.TinyIoCContainer.Current;
		private IEngine currentEngine = null;
		private Excel.Application ExcelApp;

		public string SyntaxHighlightingName { get; private set; }
		public bool CanRun { get; private set; } = false;

		public Runner(Excel.Application excel) {
			ExcelApp = excel;
		}


		public ScriptItem Script {
			get { return script; }
			set {
				script = value;
				if (script != null) {
					IoC.TryResolve(script.Type.ToString(), out currentEngine);
					CanRun = currentEngine != null;
					SyntaxHighlightingName = CanRun ? currentEngine.SyntaxHighlightingName : string.Empty;
				}
				else {
					CanRun = false;
				}
			}
		}
		private static ScriptItem script;


		public string Execute(string code) {
			var calcMode = ExcelApp.Calculation;
			if (CanRun) {
				try {
					ExcelApp.ScreenUpdating = false;
					ExcelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
					var cellCount = ((Excel.Range)ExcelApp.Selection).Cells.Count;

					var timer = new System.Diagnostics.Stopwatch();
					timer.Start();

					currentEngine.Execute(code, new XlsObject(ExcelApp));

					timer.Stop();
					return $"{cellCount} cells/{timer.Elapsed.Duration():mm\\:ss\\.fff}";
				}
				catch (Exception ex) {
					throw ex;
				}
				finally {
					ExcelApp.Calculation = calcMode;
					ExcelApp.ScreenUpdating = true;
				}
			}
			else {
				throw new Exception("Script can't be executed");
			}
		}

	}


}
