using Microsoft.ClearScript;
using Microsoft.ClearScript.Windows;
using System;

namespace ScriptAddin.Engines
{
	internal class VbEngine : EngineBase<VBScriptEngine>
	{
		public VbEngine() {
			Type = ScriptType.VbScript;
			SyntaxHighlightingName = "VB";
		}
	}
}