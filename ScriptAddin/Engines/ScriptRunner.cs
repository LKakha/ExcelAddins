using ScriptAddin.Engines;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avalon = ICSharpCode.AvalonEdit;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using System.Runtime.InteropServices.WindowsRuntime;
using ScriptAddin.Host;

namespace ScriptAddin
{
	static class ScriptRunner
	{
		static TinyIoC.TinyIoCContainer IoC = TinyIoC.TinyIoCContainer.Current;

		static ScriptRunner() {
			IoC.Register<IEngine, CSharp>("CSharp").AsSingleton();
		}

		private static IEngine currentEngine = null;
		private static Excel.Application ExcelApp = (Excel.Application)ExcelDnaUtil.Application;

		public static string SyntaxHighlightingName { get; private set; }
		public static bool CanRun { get; private set; } = false;

		public static ScriptItem Script {
			get { return script; }
			set {
				script = value;
				if (script != null) {
					IoC.TryResolve(script.Type.ToString(), out currentEngine);
					CanRun = currentEngine != null;
					SyntaxHighlightingName = CanRun ? currentEngine.SyntaxHighlightingName : string.Empty;
				} else {
					CanRun = false;
				}
			}
		}
		private static ScriptItem script;


		public static string Execute(string code) {
			var calcMode = ExcelApp.Calculation;
			if (CanRun) {
				try {
					ExcelApp.ScreenUpdating = false;
					ExcelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
					var cellCount = ((Excel.Range)ExcelApp.Selection).Cells.Count;

					var timer = new System.Diagnostics.Stopwatch();
					timer.Start();

					currentEngine.Execute(code, new HostObject(ExcelApp));

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
			} else {
				throw new Exception("Script can't be executed");
			}
		}

	}


}
