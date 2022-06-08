
namespace ScriptAddin.Engines
{
	internal interface IEngine
	{
		ScriptType Type { get; }

		string SyntaxHighlightingName { get; }

		void Execute(string code, HostObject host);

	}
}
