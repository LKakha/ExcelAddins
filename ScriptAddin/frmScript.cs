using ScriptAddin.Engines;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Avalon = ICSharpCode.AvalonEdit;
#if !DEBUG
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;
#endif

namespace ScriptAddin
{
	public partial class frmScript : Form
	{
		static readonly TinyIoC.TinyIoCContainer IoC = TinyIoC.TinyIoCContainer.Current;
		private const string caption = "Excel Scripts";
		private Db db;
		private List<ScriptItem> ScriptList;
		private Runner Runner;



		static frmScript() {
#if !DEBUG
			IoC.Register(ExcelDnaUtil.Application as Excel.Application);
#endif
			IoC.Register<Runner>().AsSingleton();
			IoC.Register<Db>().AsSingleton();
			IoC.Register<frmScript>().AsSingleton();

			IoC.Register<IEngine, CSharp>(ScriptType.CSharp.ToString()).AsSingleton();
			IoC.Register<IEngine, VbEngine>(ScriptType.VbScript.ToString()).AsSingleton();
			IoC.Register<IEngine, JsEngine>(ScriptType.JScript.ToString()).AsSingleton();
			IoC.Register<IEngine, Python>(ScriptType.Python.ToString()).AsSingleton();
		}

		public frmScript() {
			InitializeComponent();
		}

		private void frmScripts_Load(object sender, EventArgs e) {
			CodeEditor.Control.Options.IndentationSize = 2;
			db = IoC.Resolve<Db>();
			ScriptList = db.GetScriptItems().ToList();
			loadTree(null, ref ScriptList);
			Runner = IoC.Resolve<Runner>();
			CurrentScript = null;

			btnNew.DropDownItems.AddRange(
				Enum.GetNames(typeof(ScriptType)).Select(x => {
					return new ToolStripMenuItem(x, null, newScript);
				}).ToArray());
		}


		private void loadTree(TreeNode parentNode, ref List<ScriptItem> items) {
			var parentId = Guid.Empty;
			if (parentNode != null) parentId = (Guid)parentNode.Tag;
			foreach (var i in items.Where(x => x.ParentID == parentId)) {
				var node = buildNode(i);
				if (parentNode == null)
					tvScripts.Nodes.Add(node);
				else
					parentNode.Nodes.Add(node);
				loadTree(node, ref items);
			}
		}

		private TreeNode buildNode(ScriptItem item) {
			var node = new TreeNode {
				Tag = item.ID,
				Text = item.Name
			};
			if (item.Type == ScriptType.Folder)
				node.ImageIndex = 2;
			else if (item.ParentID == Guid.Empty)
				node.ImageIndex = 1;
			else
				node.ImageIndex = 0;
			node.SelectedImageIndex = node.ImageIndex;
			return node;
		}

		private ScriptItem CurrentScript {
			get { return currentScript; }
			set {
				currentScript = value;
				Runner.Script = currentScript;
				CodeEditor.SyntaxHighlighting = Avalon.Highlighting.HighlightingManager.Instance.GetDefinition(Runner.SyntaxHighlightingName ?? string.Empty);
				CodeEditor.Text = currentScript?.Code;
				if (currentScript != null && currentScript.Type != ScriptType.Folder) {
					this.Text = $"{caption} : {currentScript.Name} : {currentScript.Type}";
				}
				else { this.Text = caption; }

				if (Runner.CanRun) {
					btnRun.Enabled = true;
					CodeEditor.IsReadOnly = false;
					btnSave.Enabled = true;
					btnDelete.Enabled = true;
					this.Text = $"{caption} : {currentScript.Name} : {currentScript.Type}";
				}
				else {
					btnRun.Enabled = false;
					CodeEditor.IsReadOnly = true;
					btnSave.Enabled = false;
					btnDelete.Enabled = currentScript != null && !ScriptList.Any(x => x.ParentID == currentScript.ID);
				}
			}
		}
		private ScriptItem currentScript;


		#region Edit
		private void newScript(object sender, EventArgs e) {
			var btn = (ToolStripMenuItem)sender;
			if (Enum.TryParse<ScriptType>(btn.Text.ToString(), out var type)) {
				createNew(null, type);
			}
		}

		private void createNew(ScriptItem item, ScriptType type = ScriptType.Folder) {
			string name;
			item = (item == null) ? ScriptItem.CreateScript(type) : ScriptItem.CopyScript(item);
			name = getNewName(item.Name);
			if (name == null) return;
			item.Name = name;

			var newNode = buildNode(item);
			var currentNode = tvScripts.SelectedNode;
			if (currentNode == null) {
				tvScripts.Nodes.Add(newNode);
			}
			else {
				if (CurrentScript.Type == ScriptType.Folder) {
					currentNode.Nodes.Add(newNode);
					currentNode.Expand();
					item.ParentID = (Guid)currentNode.Tag;
				}
				else {
					if (currentNode.Parent == null) {
						tvScripts.Nodes.Add(newNode);
					}
					else {
						currentNode.Parent.Nodes.Add(newNode);
						item.ParentID = (Guid)currentNode.Parent.Tag;
					}
				}
			}

			db.Store(item);
			ScriptList.Add(item);
			CurrentScript = item;
			tvScripts.SelectedNode = newNode;
		}

		private string getNewName(string name) {
			var editor = new frmNameEditor {
				EditedString = name
			};
			if (editor.ShowDialog() == DialogResult.Cancel) return null;
			name = editor.EditedString;
			editor.Dispose();
			return name;
		}

		private void delete(object sender, EventArgs e) {
			db.Delete(CurrentScript);
			ScriptList.Remove(CurrentScript);
			tvScripts.SelectedNode.Remove();
			CurrentScript = getItemByNode(tvScripts.SelectedNode);
		}

		private void btnSave_Click(object sender, EventArgs e) {
			CurrentScript.Code = CodeEditor.Text;
			db.Store(CurrentScript);
		}

		private void btnSaveAs_Click(object sender, EventArgs e) {
			createNew(CurrentScript);
		}

		private ScriptItem getItemByNode(TreeNode node) {
			if (node == null) return null;
			var id = (Guid)node.Tag;
			return ScriptList.First(x => x.ID == id);
		}
		#endregion

		#region TreeView
		private void tvScripts_AfterSelect(object sender, TreeViewEventArgs e) {
			CurrentScript = getItemByNode(tvScripts.SelectedNode);
		}

		private void tvScripts_ItemDrag(object sender, ItemDragEventArgs e) {
			DoDragDrop(e.Item, DragDropEffects.Move);
		}

		private void tvScripts_DragEnter(object sender, DragEventArgs e) {
			e.Effect = DragDropEffects.Move;
		}

		private void tvScripts_DragDrop(object sender, DragEventArgs e) {
			var draggedNode = (TreeNode)e.Data.GetData(typeof(TreeNode));
			var targetNode = tvScripts.GetNodeAt(tvScripts.PointToClient(new System.Drawing.Point(e.X, e.Y)));

			if (!draggedNode.Equals(targetNode)) {
				draggedNode.Remove();

				ScriptItem targetItem = null;
				if (targetNode != null) {
					targetItem = getItemByNode(targetNode);
					if (targetItem.Type != ScriptType.Folder) {
						targetNode = targetNode.Parent;
						targetItem = getItemByNode(targetNode);
					}
				}

				var draggedItem = getItemByNode(draggedNode);
				if (targetNode == null) {
					tvScripts.Nodes.Add(draggedNode);
					draggedItem.ParentID = Guid.Empty;
				}
				else {
					targetNode.Nodes.Add(draggedNode);
					targetNode.Expand();
					draggedItem.ParentID = targetItem.ID;
				}
				if (draggedItem.Type == ScriptType.Folder)
					draggedNode.ImageIndex = 2;
				else if (draggedItem.ParentID == Guid.Empty)
					draggedNode.ImageIndex = 1;
				else
					draggedNode.ImageIndex = 0;
				draggedNode.SelectedImageIndex = draggedNode.ImageIndex;

				db.Store(draggedItem);
			}
		}

		private void tvScripts_AfterLabelEdit(object sender, NodeLabelEditEventArgs e) {
			var name = e.Label;
			if (!string.IsNullOrWhiteSpace(name)) {
				CurrentScript.Name = name;
				db.Store(CurrentScript);
			}
			else {
				e.CancelEdit = true;
			}
		}

		#endregion

		#region Execution

		private void btnRun_Click(object sender, EventArgs e) {
			try {
				lblStatus.Text = string.Empty;
				var result = Runner.Execute(CodeEditor.Text);
				lblStatus.Text = result;
			}
			catch (Exception ex) {
				MessageBox.Show(ex.Message);
			}
		}
		#endregion

		private void ActivateForm(object sender, EventArgs e) {
			//	this.Activate();
		}

		private void frmScript_MouseEnter(object sender, EventArgs e) {
			this.Activate();
		}
	}
}
