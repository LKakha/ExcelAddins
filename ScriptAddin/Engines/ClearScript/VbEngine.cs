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
	internal class VbEngine : IEngine
	{
		public ScriptType Type => ScriptType.VbScript;
		public string SyntaxHighlightingName { get; } = "VB";

		private const HostItemFlags flags = HostItemFlags.DirectAccess;
		private VBScriptEngine engine;

		public void Execute(string code, HostObject host) {
			try {
				using (engine = new VBScriptEngine()) {
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

