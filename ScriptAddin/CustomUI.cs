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
			return @"
<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='OnLoad' >
	<ribbon>
		<tabs>
			<tab idMso='TabAddIns'>
				<group id='ScriptAddin' label='Scripts'>
					<button id='btnOpen' onAction='btnOpen_Click' size='large' imageMso='HappyFace' showImage='true' />
				</group>
				<group id='Convertor' label='Convertor'>
					<dropDown id='ddFrom' label='From' getItemCount='GetEncodingCount' getItemLabel='GetEncodingName' onAction='encFrom_Selected' />
					<dropDown id='ddTo' label='To' getItemCount='GetEncodingCount' getItemLabel='GetEncodingName' onAction='encTo_Selected' />
					<button id='btnConvert' label='Convert' onAction='btnConvert_Click' />
				</group>
			</tab>
		</tabs>
	</ribbon>
</customUI>";
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



















		private void mainFormClosed(object sender, FormClosedEventArgs e) {
			MainForm = null;
		}
	}
}