using ExcelDna.Integration.CustomUI;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace ScriptAddin
{
	[ComVisible(true)]
	public class CustomUI : ExcelRibbon
	{
		private IRibbonUI Ribbon;

		#region CustomUI

		public override string GetCustomUI(string uiName) {
			return Properties.Resources.ui;
		}

		public void OnLoad(IRibbonUI RibbonUI) {
			try {
				Ribbon = RibbonUI;
			}
			catch (Exception ex) {
				MessageBox.Show(ex.Message);
			}
		}
		#endregion

		public void btnOpen_Click(IRibbonControl control) {
			try {
				var MainForm = new frmScript();
				MainForm.WindowState = FormWindowState.Normal;
				MainForm.Show();
			}
			catch (Exception ex) {
				MessageBox.Show(ex.Message);
			}
		}



		#region Convertor

		public int GetEncodingCount(IRibbonControl control) {
			return Convertor.GetEncodingCount();
		}

		public string GetEncodingName(IRibbonControl control, int index) {
			return Convertor.GetEncodingName(index);
		}

		private int encFrom, encTo;

		public void encFrom_Selected(IRibbonControl control, string selectedId, int selectedIndex) {
			encFrom = selectedIndex;
		}

		public void encTo_Selected(IRibbonControl control, string selectedId, int selectedIndex) {
			encTo = selectedIndex;
		}

		public void btnConvert_Click(IRibbonControl control) {
			Convertor.Convert(encFrom, encTo);
		}
		#endregion
	}
}