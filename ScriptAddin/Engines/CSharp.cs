using System;
using CSScriptLibrary;
using ScriptAddin.Host;

namespace ScriptAddin.Engines
{
	internal class CSharp : IEngine
	{
		public ScriptType Type { get; } = ScriptType.CSharp;
		public string SyntaxHighlightingName { get; } = "C#";

		public CSharp() {
			CSScript.EvaluatorConfig.Access = EvaluatorAccess.Singleton;
			CSScript.EvaluatorConfig.Engine = EvaluatorEngine.Roslyn;
		}

		public void Execute(string code, HostObject host) {
			dynamic script = CSScript.LoadCode(code).CreateObject("Script");
			script.Main(host);
		}
	}
}