using Microsoft.ClearScript;
using System;

namespace ScriptAddin.Engines
{

	internal class JsV8Engine : IEngine
	{
		public ScriptType Type => ScriptType.VbScript;
		public string SyntaxHighlightingName => "JavaScript";
		private readonly HostItemFlags flags = HostItemFlags.DirectAccess | HostItemFlags.GlobalMembers;
		private Microsoft.ClearScript.V8.V8ScriptEngine engine;

		public void Execute(string code, XlsObject xls) {
			try {
				using (engine = new Microsoft.ClearScript.V8.V8ScriptEngine()) {
					engine.AddHostObject("host", new HostFunctions());
					engine.AddHostObject("clr", new HostTypeCollection("mscorlib", "System", "System.Core"));
					engine.AddHostObject("Excel", flags, xls.App);
					engine.AddHostObject("xls", flags, xls);

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