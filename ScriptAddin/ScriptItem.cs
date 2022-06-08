using ScriptAddin.Engines;
using System;
using System.Globalization;
using System.Windows.Forms;


namespace ScriptAddin
{
	internal class ScriptItem
	{
		public ScriptItem() { }

		public Guid ID { get; set; }
		public Guid ParentID { get; set; }
		public string Name { get; set; }
		public string Code { get; set; }
		public ScriptType Type { get; set; }

		public static ScriptItem CreateScript(ScriptType type) {
			var name = type == ScriptType.Folder ? "New Folder" : $"New {type} Script";
			return new ScriptItem { ID = Guid.NewGuid(), Name = name, Type = type };
		}

		public static ScriptItem CopyScript(ScriptItem item) {
			return new ScriptItem {
				ID = Guid.NewGuid(),
				ParentID = item.ParentID,
				Name = $"Copy of {item.Name}",
				Code = item.Code,
				Type = item.Type
			};
		}
	}
}
