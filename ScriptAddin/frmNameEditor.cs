using System;
using System.Windows.Forms;

namespace ScriptAddin
{
	public partial class frmNameEditor : Form
	{
		public frmNameEditor() {
			InitializeComponent();
		}

		public string EditedString {
			get { return TextBox1.Text; }
			set { TextBox1.Text = value; }
		}

		private void btnOK_Click(object sender, EventArgs e) {
			DialogResult = DialogResult.OK;
		}

	}
}
