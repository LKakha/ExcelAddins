using IronPython.Hosting;
using Microsoft.Scripting.Hosting;
using System;

namespace ScriptAddin.Engines
{
	internal class Python : IEngine
	{
		public ScriptType Type => ScriptType.Python;
		public string SyntaxHighlightingName => "Python";

		private ScriptEngine engine;

		public Python() {
			engine = IronPython.Hosting.Python.CreateEngine();
		}

		public void Execute(string code, XlsObject xls) {
			try {
				ScriptScope scope = engine.CreateScope();
				scope.ImportModule("clr");
				ScriptSource source = engine.CreateScriptSourceFromString(code);
				scope.SetVariable("xls", xls);
				source.Execute(scope);
			}
			catch (Exception ex) {
				throw ex;
			}
		}
	}
}
