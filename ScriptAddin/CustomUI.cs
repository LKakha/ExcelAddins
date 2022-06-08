using ExcelDna.Integration.CustomUI;
using ScriptAddin.Engines;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ScriptAddin
{
	[ComVisible(true)]
	public class CustomUI : ExcelRibbon
	{
		private Microsoft.Office.Interop.Excel.Application Excel;
		private IRibbonUI Ribbon;
		private frmScript MainForm;

		#region CustomUI

		public override string GetCustomUI(string uiName) {
			return Properties.Resources.ui;
		}

		public void OnLoad(IRibbonUI RibbonUI) {
			try {
				Excel = (Microsoft.Office.Interop.Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
				Ribbon = RibbonUI;
			}
			catch (Exception ex) {
				MessageBox.Show(ex.Message);
			}
		}
		#endregion

		public void btnOpen_Click(IRibbonControl control) {
			if (MainForm == null) {
				try {
					MainForm = new frmScript();
					MainForm.FormClosed += mainFormClosed;
				}
				catch (Exception ex) {
					MessageBox.Show(ex.Message);
				}
			}
			MainForm.WindowState = FormWindowState.Normal;
			MainForm.Show();
		}

		private void mainFormClosed(object sender, FormClosedEventArgs e) {
			MainForm = null;
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