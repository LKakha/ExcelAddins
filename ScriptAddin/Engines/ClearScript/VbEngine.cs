using Microsoft.ClearScript;
using Microsoft.ClearScript.Windows;
using System;

namespace ScriptAddin.Engines
{
	internal class VbEngine : IEngine
	{
		public ScriptType Type => ScriptType.VbScript;
		public string SyntaxHighlightingName => "VB";
		private readonly HostItemFlags flags = HostItemFlags.DirectAccess | HostItemFlags.GlobalMembers;
		private VBScriptEngine engine;

		public void Execute(string code, XlsObject xls) {
			try {
				using (engine = new VBScriptEngine()) {
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