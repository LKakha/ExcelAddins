using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ScriptAddin.Engines
{
	public class ScriptExtension : Microsoft.ClearScript.ExtendedHostFunctions
	{
		public DialogResult MsgBox(object Msg, MessageBoxButtons Button = MessageBoxButtons.OK) {
			return MessageBox.Show(Msg?.ToString(), "ScriptAddin", Button);
		}
		public DialogResult alert(object Msg, MessageBoxButtons Button = MessageBoxButtons.OK) {
			return MsgBox(Msg, Button);
		}
	}
}
