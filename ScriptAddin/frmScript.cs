using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Avalon = ICSharpCode.AvalonEdit;
using Excel = Microsoft.Office.Interop.Excel;

namespace ScriptAddin
{
	public partial class frmScript : Form
	{
		public frmScript() {
			try {
				InitializeComponent();
			}
			catch (Exception ex) {
				MessageBox.Show(ex.Message);
			}
		}

		private const string caption = "Excel Scripts";
		private Db4objects.Db4o.IObjectContainer db;
		private List<ScriptItem> ScriptList;

		private void frmScripts_Load(object sender, EventArgs e) {
			CodeEditor.Control.Options.IndentationSize = 2;
			getEngineTypes();
			loadScripts();
			loadTree(null, ref ScriptList);
			CurrentScript = null;
		}
		private void getEngineTypes() {
			var engineTypes = ScriptRunner.SupportedEngines;
			foreach (var e in engineTypes) {
				btnNew.DropDownItems.Add(new ToolStripMenuItem($"New {e}", null, btnNew_Click) {
					Tag = e
				}); ;
			}
		}
		private void loadScripts() {
			db = Db4objects.Db4o.Db4oEmbedded.OpenFile(System.IO.Path.GetDirectoryName(ExcelDnaUtil.XllPath) + "\\Scripts.db");
			//db = Db4objects.Db4o.Db4oEmbedded.OpenFile("Scripts.db");
			ScriptList = (from i in db.Query<ScriptItem>() orderby i.Type, i.Name select i).ToList();
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
			var node = new TreeNode();
			node.Tag = item.ID;
			node.Text = item.Name;
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
				if (currentScript == null || currentScript.Type == ScriptType.Folder) {
					CodeEditor.IsEnabled = false;
					CodeEditor.Text = null;
					btnRun.Enabled = false;
					btnSave.Enabled = false;
					if (currentScript == null)
						btnDelete.Enabled = false;
					else {
						btnDelete.Enabled = !ScriptList.Any(x => x.ParentID == currentScript.ID);
					}
					this.Text = caption;
				} else {
					ScriptRunner.Script = currentScript;
					CodeEditor.SyntaxHighlighting = ScriptRunner.SyntaxHighlighting;

					this.Text = $"{caption} : {currentScript.Name} : {currentScript.Type}";
					CodeEditor.Text = currentScript.Code;

					btnRun.Enabled = true;
					CodeEditor.IsEnabled = true;
					btnSave.Enabled = true;
					btnDelete.Enabled = true;
				}
			}
		}
		private ScriptItem currentScript;

		private void btnNew_Click(object sender, EventArgs e) {
			var btn = (ToolStripMenuItem)sender;
			if (btn.Tag is ScriptType type) {
				createNew(null, type);
			}
		}

		private void createNew(ScriptItem item, ScriptType type = ScriptType.Folder) {
			string name;
			if (item == null) {
				item = ScriptItem.CreateScript(type);
			} else {
				item = ScriptItem.CopyScript(item);
			}
			name = getNewName(item.Name);
			if (name == null) return;
			item.Name = name;

			var newNode = buildNode(item);
			var currentNode = tvScripts.SelectedNode;
			if (currentNode == null) {
				tvScripts.Nodes.Add(newNode);
			} else {
				if (CurrentScript.Type == ScriptType.Folder) {
					currentNode.Nodes.Add(newNode);
					currentNode.Expand();
					item.ParentID = (Guid)currentNode.Tag;
				} else {
					if (currentNode.Parent == null) {
						tvScripts.Nodes.Add(newNode);
					} else {
						currentNode.Parent.Nodes.Add(newNode);
						item.ParentID = (Guid)currentNode.Parent.Tag;
					}
				}
			}

			db.Store(item);
			db.Commit();
			ScriptList.Add(item);
			CurrentScript = item;
			tvScripts.SelectedNode = newNode;
		}

		private string getNewName(string name) {
			var editor = new frmEditor();
			editor.EditedString = name;
			if (editor.ShowDialog() == DialogResult.Cancel) return null;
			name = editor.EditedString;
			editor.Dispose();
			return name;
		}

		private void delete(object sender, EventArgs e) {
			db.Delete(CurrentScript);
			db.Commit();
			ScriptList.Remove(CurrentScript);
			tvScripts.SelectedNode.Remove();
			CurrentScript = getItemByNode(tvScripts.SelectedNode);
		}

		private void btnSave_Click(object sender, EventArgs e) {
			CurrentScript.Code = CodeEditor.Text;
			db.Store(CurrentScript);
			db.Commit();
		}

		private void btnSaveAs_Click(object sender, EventArgs e) {
			createNew(CurrentScript);
		}

		private ScriptItem getItemByNode(TreeNode node) {
			if (node == null) return null;
			var id = (Guid)node.Tag;
			return ScriptList.First(x => x.ID == id);
		}

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
				} else {
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
				db.Commit();
			}
		}

		private void tvScripts_AfterLabelEdit(object sender, NodeLabelEditEventArgs e) {
			var name = e.Label;
			if (!string.IsNullOrWhiteSpace(name)) {
				CurrentScript.Name = name;
				db.Store(CurrentScript);
				db.Commit();
			} else {
				e.CancelEdit = true;
			}
		}

		#endregion

		#region Execution

		private void btnRun_Click(object sender, EventArgs e) {
			try {
				lblStatus.Text = string.Empty;
				var result = ScriptRunner.Execute(CodeEditor.Text);
				lblStatus.Text = result;
			}
			catch (Exception ex) {
				MessageBox.Show(ex.Message);
			}
		}
		#endregion

		private void Me_FormClosing(object sender, FormClosingEventArgs e) {
			db.Close();
		}

		private void ActivateForm(object sender, EventArgs e) {
			this.Activate();
		}
	}
}
