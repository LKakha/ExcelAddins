using ICSharpCode.AvalonEdit;
using ScriptAddin.Host;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Markup;
using Avalon = ICSharpCode.AvalonEdit;


namespace ScriptAddin.Engines
{
	internal interface IEngine
	{
		ScriptType Type { get; }
		string SyntaxHighlightingName { get; }

		void Execute(string code, HostObject host);

	}
}
