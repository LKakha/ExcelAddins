using System;
using System.Collections.Generic;
using System.Linq;
#if !DEBUG
using ExcelDna.Integration;
#endif


namespace ScriptAddin
{
	internal class Db : IDisposable
	{
		private Db4objects.Db4o.IObjectContainer db;

		public Db() {
#if DEBUG
			db = Db4objects.Db4o.Db4oEmbedded.OpenFile(Application.StartupPath + "\\Scripts.db");
#else
			db = Db4objects.Db4o.Db4oEmbedded.OpenFile(System.IO.Path.GetDirectoryName(ExcelDnaUtil.XllPath) + "\\Scripts.db");
#endif
		}


		public IEnumerable<ScriptItem> GetScriptItems() {
			return from i in db.Query<ScriptItem>() orderby i.Type, i.Name select i;
		}

		public void Store(ScriptItem item) {
			db.Store(item);
			db.Commit();
		}

		public void Delete(ScriptItem item) {
			db.Delete(item);
			db.Commit();
		}

		public void Dispose() {
			db.Close();
			Console.WriteLine("Db closed");
		}
	}
}
