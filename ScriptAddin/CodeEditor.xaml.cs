using System.Windows.Controls;

namespace ScriptAddin
{

	public partial class CodeEditor : UserControl
	{
		public CodeEditor() {
			InitializeComponent();
		}

		public ICSharpCode.AvalonEdit.TextEditor Control
		{
			get { return Editor; }
		}

		public ICSharpCode.AvalonEdit.Highlighting.IHighlightingDefinition SyntaxHighlighting
		{
			get { return Editor.SyntaxHighlighting; }
			set { Editor.SyntaxHighlighting = value; }
		}

		public bool IsReadOnly {
			get { return Editor.IsReadOnly; }
			set { Editor.IsReadOnly = value; }
		}

		public string Text
		{
			get { return Editor.Text; }
			set { Editor.Text = value; }
		}
	}
}
