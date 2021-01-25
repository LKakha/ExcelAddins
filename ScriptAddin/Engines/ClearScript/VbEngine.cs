using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Avalon = ICSharpCode.AvalonEdit;
using Microsoft.ClearScript.Windows;
using Microsoft.ClearScript;

namespace ScriptAddin.Engines
{
	public class VbEngine : IEngine
	{
		private static Avalon.Highlighting.IHighlightingDefinition highlightingDefinition = Avalon.Highlighting.HighlightingManager.Instance.GetDefinition("VB");
		private const HostItemFlags flags = HostItemFlags.DirectAccess;
		public ScriptType Type => ScriptType.VB;
		public Avalon.Highlighting.IHighlightingDefinition HighlightingDefinition => highlightingDefinition;

		private VBScriptEngine engine;

		public void Execute(string code, Action<IEngine> initAction = null) {
			try {
				using (engine = new VBScriptEngine()) {
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

		private void initEngine(VBScriptEngine engine) {
			engine.AddHostObject("host", new HostFunctions());
			engine.AddHostObject("ext", new ScriptExtension());
			engine.AddHostObject("clr", new HostTypeCollection("mscorlib", "System", "System.Core"));
		}

		public void AddHostObject(string name, object obj) {
			engine?.AddHostObject(name, flags, obj);
		}

	}

}

