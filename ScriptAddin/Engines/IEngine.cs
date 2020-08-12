using ICSharpCode.AvalonEdit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Markup;
using Avalon = ICSharpCode.AvalonEdit;


namespace ScriptAddin.Engines
{
	interface IEngine
	{
		ScriptType Type { get; }
		Avalon.Highlighting.IHighlightingDefinition HighlightingDefinition { get; }

		void AddHostObject(string name, object obj);
		void Execute(string code, Action<IEngine> initAction = null);

	}
}
