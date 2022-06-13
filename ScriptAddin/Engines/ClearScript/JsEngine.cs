using Microsoft.ClearScript;
using Microsoft.ClearScript.Windows;
using System;

namespace ScriptAddin.Engines
{
	internal class JsEngine : EngineBase<JScriptEngine>
	{
		public JsEngine() {
			Type = ScriptType.JScript;
			SyntaxHighlightingName = "JavaScript";
		}
	}
}