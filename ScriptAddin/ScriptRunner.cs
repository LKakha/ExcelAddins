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
	static class ScriptRunner
	{

		private static readonly Dictionary<ScriptType, IEngine> engines = new Dictionary<ScriptType, IEngine>();
		private static IEngine currentEngine = null;
		private static Excel.Application ExcelApp = (Excel.Application)ExcelDnaUtil.Application;

		public static void AddEngine(IEngine engine) {
			if (!engines.ContainsKey(engine.Type)) {
				engines.Add(engine.Type, engine);
			} else {
				throw new Exception("Engine is already installed");
			}
		}

		public static ScriptType[] SupportedEngines  => engines.Keys.ToArray();

		public static Avalon.Highlighting.IHighlightingDefinition SyntaxHighlighting => currentEngine?.HighlightingDefinition;

		private static ScriptItem script;
		public static ScriptItem Script {
			get { return script; }
			set {
				script = value;
				if (script != null && engines.ContainsKey(script.Type)) {
					currentEngine = engines[script.Type];
				} else {
					currentEngine = null;
				}
			}
		}

		public static string Execute(string code) {
			var calcMode = ExcelApp.Calculation;
			if (currentEngine != null) {
				try {
					ExcelApp.ScreenUpdating = false;
					ExcelApp.Calculation = Excel.XlCalculation.xlCalculationManual;

					var timer = new System.Diagnostics.Stopwatch();
					timer.Start();

					currentEngine.Execute(code, engine => {
						engine.AddHostObject("Excel", ExcelApp);
						engine.AddHostObject("Book", ExcelApp.ActiveWorkbook);
						var sheet = ExcelApp.ActiveSheet;
						engine.AddHostObject("Sheet", sheet);
						engine.AddHostObject("Sel", (Excel.Range)ExcelApp.Selection);
					});

					timer.Stop();

					var cellCount = ((Excel.Range)ExcelApp.Selection).Cells.Count;
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
				throw new Exception($"No script engine found for script type {script.Type}");
			}
		}

	}


}
