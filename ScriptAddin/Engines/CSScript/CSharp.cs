using CSScriptLibrary;

namespace ScriptAddin.Engines
{
	internal class CSharp : IEngine
	{
		public ScriptType Type => ScriptType.CSharp;
		public string SyntaxHighlightingName => "C#";

		public CSharp() {
			CSScript.EvaluatorConfig.Access = EvaluatorAccess.Singleton;
			CSScript.EvaluatorConfig.Engine = EvaluatorEngine.Roslyn;
		}

		public void Execute(string code, XlsObject xls) {
			dynamic script = CSScript.LoadCode(code).CreateObject("Script");
			script.Main(xls);
		}
	}
}