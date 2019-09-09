using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ScriptAddin
{
	public static class Convertor
	{
		private static Dictionary<string, short[]> Encodings = new Dictionary<string, short[]>() {
				{"GEO_UTF", new short[] { 4304, 4305, 4306, 4307, 4308, 4309, 4310, 4311, 4312, 4313, 4314, 4315, 4316, 4317, 4318, 4319, 4320, 4321, 4322, 4323, 4324, 4325, 4326, 4327, 4328, 4329, 4330, 4331, 4332, 4333, 4334, 4335, 4336, 46, 47, 34} },
				{"GEO_LAT", new short[] { 97, 98, 103, 100, 101, 118, 122, 84, 105, 107, 108, 109, 110, 111, 112, 74, 114, 115, 116, 117, 102, 113, 82, 121, 83, 67, 99, 90, 119, 87, 120, 106, 104} },
				{"GEO_SGGG", new short[] { 102, 44, 117, 108, 116, 100, 112, 115, 98, 114, 107, 118, 121, 106, 103, 59, 104, 99, 110, 101, 97, 109, 113, 39, 105, 120, 119, 46, 111, 122, 91, 47, 93, 62, 63, 35} },
				{"GEO_STS", new short[] { 192, 193, 194, 195, 196, 197, 198, 199, 200, 201, 202, 203, 204, 205, 206, 207, 208, 209, 210, 211, 212, 213, 214, 215, 216, 217, 218, 219, 220, 221, 222, 223, 224, 46, 47, 34} },
				{"GEO_ABC", new short[] { 192, 193, 194, 195, 196, 197, 198, 200, 201, 202, 203, 204, 205, 207, 208, 209, 210, 211, 212, 214, 215, 216, 217, 218, 219, 220, 221, 222, 223, 224, 225, 227, 228, 106, 104} }
			};

		private static string[] EncodingNames;

		static Convertor() {
			EncodingNames = Encodings.Select(x => x.Key).ToArray();
		}

		public static int GetEncodingCount() {
			return Encodings.Count;
		}

		public static string GetEncodingName(int index) {
			return EncodingNames[index];
		}

		private static string convertInternal(string text, string from, string to) {
			var newText = new System.Text.StringBuilder(text, text.Length);
			var encSource = Encodings[from];
			var encDest = Encodings[to];

			for (var i = 0; i < Math.Min(encSource.Length, encDest.Length); i++) {
				newText.Replace(System.Convert.ToChar(encSource[i]), System.Convert.ToChar(encDest[i]));
			}
			return newText.ToString();
		}

		public static void Convert(int from, int to) {
			if (from == to) return;
			var eFrom = EncodingNames[from];
			var eTo = EncodingNames[to];
			var app = (Excel.Application)ExcelDnaUtil.Application;
			var sel = (Excel.Range)app.Selection;

			var calcMode = app.Calculation;
			app.ScreenUpdating = false;
			app.Calculation = Excel.XlCalculation.xlCalculationManual;

			try {

				for (int areaN = 1; areaN <= sel.Areas.Count; areaN++) {
					var area = sel.Areas[areaN];
					var areaCol = area.Column;
					var areaRow = area.Row;
					var areaCol2 = areaCol + area.Columns.Count - 1;
					var areaRow2 = areaRow + area.Rows.Count - 1;

					for (int col = areaCol; col <= areaCol2; col++) {
						for (int row = areaRow; row <= areaRow2; row++) {
							try {
								var cellValue = ((Excel.Range)app.Cells[row, col]).Value;
								if (cellValue != null) {
									var value = cellValue.ToString();
									if (!string.IsNullOrWhiteSpace(value)) {
										value = convertInternal(value, eFrom, eTo);
										((Excel.Range)app.Cells[row, col]).Value = value;
									}
								}
							}
							catch (Exception ex) {
								((Excel.Range)app.Cells[row, col]).Value = ex.Message;
							}
						}
					}
				}
			}
			catch (Exception ex) {
				throw ex;
			}
			finally {
				app.ScreenUpdating = true;
				app.Calculation = calcMode;
				app = null;
			}
		}
	}
}
