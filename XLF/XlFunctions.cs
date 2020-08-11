using ExcelDna.Integration;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Text;

public static class XlFunctions
{
	private static readonly string[][] GeoNum = {
		new [] {null,"ერთი","ორი","სამი","ოთხი","ხუთი","ექვსი","შვიდი","რვა","ცხრა","ათი","თერთმეტი","თორმეტი","ცამეტი","თოთხმეტი","თხუთმეტი","თექვსმეტი","ჩვიდმეტი","თვრამეტი","ცხრამეტი" },
		new [] {null,"ოც","ორმოც","სამოც","ოთხმოც" },
		new [] {null, "ას","ორას","სამას","ოთხას","ხუთას","ექვსას","შვიდას","რვაას","ცხრაას" },
		new [] {null,"ათას","მილიონ","მილიარდ","ტრილიონ","კვადრილიონ","კვინტილიონ" }
	};

	[ExcelFunction(Description = "მთელი რიცხვი ტექსტურად")]
	public static string NumToString(long num) {
		if (num == 0) return "ნული";
		var alltxt = new StringBuilder();
		var txt = new StringBuilder();
		const string ი = "ი";
		const string space = " ";

		var i = 0;

		while (num > 0) {
			var t = (short)(num % 1000);
			num /= 1000;

			if (t > 0) {
				txt.Clear();
				int s3 = t / 100;
				int s2 = (t % 100) / 20;
				int s1 = t % 20;

				txt.Append(GeoNum[2][s3]);
				if (s2 > 0) txt.Append(GeoNum[1][s2]);
				if (s1 > 0) txt.Append((s2 > 0) ? "და" : null).Append(GeoNum[0][s1]); else txt.Append(ი);
				if (i > 0) txt.Append(space).Append(GeoNum[3][i]).Append(space);
				alltxt.Insert(0, txt);
			}
			i++;
		}
		if (alltxt[alltxt.Length - 1] != ი[0]) alltxt.Replace(space, ი, alltxt.Length - 1, 1);
		return alltxt.ToString();
	}

	[ExcelFunction(Description = "თანხა ტექსტურად")]
	public static string GeoMoney(decimal num, string format = null) {
		long ლარი = (long)num;
		short თეთრი = (short)((num - ლარი) * 100);

		if (!string.IsNullOrWhiteSpace(format)) {
			var ret = format;
			ret = ret.Replace("{L}", NumToString(ლარი));
			ret = ret.Replace("{T}", NumToString((long)თეთრი));
			ret = ret.Replace("{l}", ლარი.ToString());
			ret = ret.Replace("{t}", თეთრი.ToString("D2"));
			return ret;
		}
		return $"{NumToString(ლარი)} ლარი {NumToString((long)თეთრი)} თეთრი";
	}

	[ExcelFunction]
	public static object Switch(object value, object val1, object ret1, object val2, object ret2, object val3, object ret3, object val4, object ret4, object val5, object ret5) {
		if (Equals(value, val1)) return ret1;
		if (Equals(value, val2)) return ret2;
		if (Equals(value, val3)) return ret3;
		if (Equals(value, val4)) return ret4;
		if (Equals(value, val5)) return ret5;

		return null;
	}

}
