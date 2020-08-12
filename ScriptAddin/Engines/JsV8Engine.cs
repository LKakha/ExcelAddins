using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Avalon = ICSharpCode.AvalonEdit;
using Microsoft.ClearScript;
using Microsoft.ClearScript.V8;

namespace ScriptAddin.Engines
{
	class JsV8Engine : IEngine
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

					engine.Execute(code);
					engine.CollectGarbage(false);
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
			engine.AllowReflection = true;
			engine.AddHostObject("host", flags, new HostFunctions());
			engine.AddHostObject("ext", flags, new ScriptExtension());
			engine.AddHostObject("clr", flags, new HostTypeCollection("mscorlib", "System", "System.Core"));
		}

		public void AddHostObject(string name, object obj) {
			engine?.AddHostObject(name, flags, obj);
		}
	}
}