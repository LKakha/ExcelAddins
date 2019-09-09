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

	[ExcelFunction(Description = "რიცხვი ტექსტურად")]
	public static string NumToString(ulong num) {
		var triada = new short[GeoNum[3].Length];
		var i = 0;
		var alltxt = new StringBuilder();
		var txt = new StringBuilder();
		const string ი = "ი";
		const string space = " ";

		while (num > 0) {
			triada[i] = (short)(num % 1000);
			num /= 1000;
			i++;
		}

		for (i = triada.Length - 1; i >= 0; i--) {
			var t = triada[i];
			if (t > 0) {
				txt.Clear();
				int s3 = t / 100;
				int s2 = (t % 100) / 20;
				int s1 = t % 20;

				txt.Append(GeoNum[2][s3]);
				if (s2 > 0) txt.Append(GeoNum[1][s2]);
				if (s1 > 0) txt.Append((s2 > 0) ? "და" : null).Append(GeoNum[0][s1]); else txt.Append(ი);
				if (i > 0) txt.Append(space).Append(GeoNum[3][i]).Append(space);

				alltxt.Append(txt);
			}
		}
		if (alltxt[alltxt.Length - 1] != ი[0]) alltxt.Replace(space, ი, alltxt.Length - 1, 1);
		return alltxt.ToString();
	}
}
