using ExcelDna.Integration.CustomUI;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace TIDB
{
	[ComVisible(true)]
	public class CustomUI : ExcelRibbon
	{
		private Microsoft.Office.Interop.Excel.Application Excel;
		private IRibbonUI Ribbon;

		#region CustomUI

		public override string GetCustomUI(string uiName) {
			return @"
<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='OnLoad' >
	<ribbon>
		<tabs>
			<tab idMso='TabAddIns'>
				<group id='TIDB' label='TIDB'>
					<button id='btnConnect' label='Connect' onAction='btnConnect_Click' />
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
				MyFunctions.SetConnection();
			}
			catch (Exception ex) {
				MessageBox.Show(ex.Message);
			}
		}

		public void btnConnect_Click(IRibbonControl control) {
			MyFunctions.SetConnection();
		}
		#endregion
	}
}