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

	internal class JsV8Engine : IEngine
	{
		public ScriptType Type => ScriptType.VbScript;
		public string SyntaxHighlightingName { get; } = "JavaScript";

		private Microsoft.ClearScript.V8.V8ScriptEngine engine;

		public void Execute(string code, HostObject host) {
			try {
				using (engine = new Microsoft.ClearScript.V8.V8ScriptEngine()) {
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
				throw ex;
			}
		}
	}
}