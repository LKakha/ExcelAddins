using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Avalon = ICSharpCode.AvalonEdit;
using Microsoft.ClearScript;
using Microsoft.ClearScript.V8;
using ExcelDna.Integration;

namespace ScriptAddin.Engines
{
	public class JsV8Engine : IEngine
	{
		private static Avalon.Highlighting.IHighlightingDefinition highlightingDefinition = Avalon.Highlighting.HighlightingManager.Instance.GetDefinition("JavaScript");
		private const HostItemFlags flags = HostItemFlags.DirectAccess;
		public ScriptType Type => ScriptType.JSV8;
		public Avalon.Highlighting.IHighlightingDefinition HighlightingDefinition => highlightingDefinition;

		private V8ScriptEngine engine;

		public void Execute(string code, Action<IEngine> initAction = null) {
			try {
				using (engine = new V8ScriptEngine()) {
					initEngine(engine);
					initAction(this);

					var bin = engine.Compile(code);
					engine.Execute(bin);
					engine.CollectGarbage(true);
				}
			}
			catch (ScriptEngineException ex) {
				throw new Exception(ex.ErrorDetails);
			}
			catch (Exception ex) {
				throw ex;
			}
		}

		private void initEngine(V8ScriptEngine engine) {
			engine.AddHostObject("host", flags, new HostFunctions());
			engine.AddHostObject("ext", flags, new ScriptExtension());
			engine.AddHostObject("clr", flags, new HostTypeCollection("mscorlib", "System", "System.Core"));
		}

		public void AddHostObject(string name, object obj) {
			//var a=ExcelDna.Integration.			engine?.AddHostObject(name, flags, obj);
		}
	}
}