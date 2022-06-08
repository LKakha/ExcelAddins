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

namespace ScriptAddin
{
	static class Runner
	{
		static readonly TinyIoC.TinyIoCContainer IoC = TinyIoC.TinyIoCContainer.Current;

		static Runner() {
			IoC.Register<IEngine, CSharp>(ScriptType.CSharp.ToString()).AsSingleton();
			IoC.Register<IEngine, VbEngine>(ScriptType.VbScript.ToString()).AsSingleton();
			IoC.Register<IEngine, JsEngine>(ScriptType.JScript.ToString()).AsSingleton();
			IoC.Register<IEngine, JsV8Engine>(ScriptType.JsV8.ToString()).AsSingleton();
		}

		private static IEngine currentEngine = null;
		private static Excel.Application ExcelApp = null;
#if !DEBUG
		ExcelApp= (Excel.Application) ExcelDnaUtil.Application;
#endif

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
				}
				else {
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
			}
			else {
				throw new Exception("Script can't be executed");
			}
		}

	}


}
