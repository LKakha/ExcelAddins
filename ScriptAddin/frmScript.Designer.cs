namespace ScriptAddin
{
	partial class frmScript
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing) {
			if (disposing && (components != null)) {
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent() {
			this.components = new System.ComponentModel.Container();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmScript));
			this.toolStrip1 = new System.Windows.Forms.ToolStrip();
			this.btnRun = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
			this.btnNew = new System.Windows.Forms.ToolStripDropDownButton();
			this.btnJS = new System.Windows.Forms.ToolStripMenuItem();
			this.btnVB = new System.Windows.Forms.ToolStripMenuItem();
			this.btnFolder = new System.Windows.Forms.ToolStripMenuItem();
			this.btnSave = new System.Windows.Forms.ToolStripSplitButton();
			this.btnSaveAs = new System.Windows.Forms.ToolStripMenuItem();
			this.btnDelete = new System.Windows.Forms.ToolStripButton();
			this.btnHelp = new System.Windows.Forms.ToolStripButton();
			this.lblStatus = new System.Windows.Forms.ToolStripLabel();
			this.splitContainer1 = new System.Windows.Forms.SplitContainer();
			this.tvScripts = new System.Windows.Forms.TreeView();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.CodeEditorHost = new System.Windows.Forms.Integration.ElementHost();
			this.CodeEditor = new ScriptAddin.CodeEditor();
			this.btnJSV8 = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStrip1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
			this.splitContainer1.Panel1.SuspendLayout();
			this.splitContainer1.Panel2.SuspendLayout();
			this.splitContainer1.SuspendLayout();
			this.SuspendLayout();
			// 
			// toolStrip1
			// 
			this.toolStrip1.BackColor = System.Drawing.Color.LightGray;
			this.toolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.toolStrip1.ImageScalingSize = new System.Drawing.Size(24, 24);
			this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnRun,
            this.toolStripSeparator1,
            this.btnNew,
            this.btnSave,
            this.btnDelete,
            this.btnHelp,
            this.lblStatus});
			this.toolStrip1.Location = new System.Drawing.Point(0, 419);
			this.toolStrip1.Name = "toolStrip1";
			this.toolStrip1.Size = new System.Drawing.Size(800, 31);
			this.toolStrip1.TabIndex = 0;
			// 
			// btnRun
			// 
			this.btnRun.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
			this.btnRun.Image = ((System.Drawing.Image)(resources.GetObject("btnRun.Image")));
			this.btnRun.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnRun.Name = "btnRun";
			this.btnRun.Size = new System.Drawing.Size(28, 28);
			this.btnRun.Text = "Run";
			this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
			this.btnRun.MouseEnter += new System.EventHandler(this.ActivateForm);
			// 
			// toolStripSeparator1
			// 
			this.toolStripSeparator1.Name = "toolStripSeparator1";
			this.toolStripSeparator1.Size = new System.Drawing.Size(6, 31);
			// 
			// btnNew
			// 
			this.btnNew.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
			this.btnNew.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnJS,
            this.btnVB,
            this.btnFolder,
            this.btnJSV8});
			this.btnNew.Image = ((System.Drawing.Image)(resources.GetObject("btnNew.Image")));
			this.btnNew.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnNew.Name = "btnNew";
			this.btnNew.Size = new System.Drawing.Size(37, 28);
			this.btnNew.Text = "Create New";
			// 
			// btnJS
			// 
			this.btnJS.Image = ((System.Drawing.Image)(resources.GetObject("btnJS.Image")));
			this.btnJS.Name = "btnJS";
			this.btnJS.Size = new System.Drawing.Size(188, 30);
			this.btnJS.Text = "New JScript";
			this.btnJS.Click += new System.EventHandler(this.btnNew_Click);
			// 
			// btnVB
			// 
			this.btnVB.Image = ((System.Drawing.Image)(resources.GetObject("btnVB.Image")));
			this.btnVB.Name = "btnVB";
			this.btnVB.Size = new System.Drawing.Size(188, 30);
			this.btnVB.Text = "New VBScript";
			this.btnVB.Click += new System.EventHandler(this.btnNew_Click);
			// 
			// btnFolder
			// 
			this.btnFolder.Image = ((System.Drawing.Image)(resources.GetObject("btnFolder.Image")));
			this.btnFolder.Name = "btnFolder";
			this.btnFolder.Size = new System.Drawing.Size(188, 30);
			this.btnFolder.Text = "New Folder";
			this.btnFolder.Click += new System.EventHandler(this.btnNew_Click);
			// 
			// btnSave
			// 
			this.btnSave.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
			this.btnSave.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnSaveAs});
			this.btnSave.Image = ((System.Drawing.Image)(resources.GetObject("btnSave.Image")));
			this.btnSave.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnSave.Name = "btnSave";
			this.btnSave.Size = new System.Drawing.Size(40, 28);
			this.btnSave.ToolTipText = "Save";
			this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
			// 
			// btnSaveAs
			// 
			this.btnSaveAs.Image = ((System.Drawing.Image)(resources.GetObject("btnSaveAs.Image")));
			this.btnSaveAs.Name = "btnSaveAs";
			this.btnSaveAs.Size = new System.Drawing.Size(114, 22);
			this.btnSaveAs.Text = "Save As";
			this.btnSaveAs.Click += new System.EventHandler(this.btnSaveAs_Click);
			// 
			// btnDelete
			// 
			this.btnDelete.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
			this.btnDelete.Image = ((System.Drawing.Image)(resources.GetObject("btnDelete.Image")));
			this.btnDelete.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnDelete.Name = "btnDelete";
			this.btnDelete.Size = new System.Drawing.Size(28, 28);
			this.btnDelete.Text = "Delete";
			this.btnDelete.Click += new System.EventHandler(this.delete);
			// 
			// btnHelp
			// 
			this.btnHelp.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
			this.btnHelp.Image = ((System.Drawing.Image)(resources.GetObject("btnHelp.Image")));
			this.btnHelp.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.btnHelp.Name = "btnHelp";
			this.btnHelp.Size = new System.Drawing.Size(28, 28);
			this.btnHelp.Text = "Help";
			// 
			// lblStatus
			// 
			this.lblStatus.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
			this.lblStatus.AutoSize = false;
			this.lblStatus.Name = "lblStatus";
			this.lblStatus.Size = new System.Drawing.Size(150, 28);
			// 
			// splitContainer1
			// 
			this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.splitContainer1.Location = new System.Drawing.Point(0, 0);
			this.splitContainer1.Name = "splitContainer1";
			// 
			// splitContainer1.Panel1
			// 
			this.splitContainer1.Panel1.Controls.Add(this.tvScripts);
			// 
			// splitContainer1.Panel2
			// 
			this.splitContainer1.Panel2.Controls.Add(this.CodeEditorHost);
			this.splitContainer1.Size = new System.Drawing.Size(800, 419);
			this.splitContainer1.SplitterDistance = 266;
			this.splitContainer1.TabIndex = 2;
			// 
			// tvScripts
			// 
			this.tvScripts.AllowDrop = true;
			this.tvScripts.BackColor = System.Drawing.Color.LightGray;
			this.tvScripts.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tvScripts.FullRowSelect = true;
			this.tvScripts.HideSelection = false;
			this.tvScripts.ImageIndex = 0;
			this.tvScripts.ImageList = this.imageList1;
			this.tvScripts.LabelEdit = true;
			this.tvScripts.Location = new System.Drawing.Point(0, 0);
			this.tvScripts.Name = "tvScripts";
			this.tvScripts.SelectedImageIndex = 0;
			this.tvScripts.ShowRootLines = false;
			this.tvScripts.Size = new System.Drawing.Size(266, 419);
			this.tvScripts.TabIndex = 0;
			this.tvScripts.AfterLabelEdit += new System.Windows.Forms.NodeLabelEditEventHandler(this.tvScripts_AfterLabelEdit);
			this.tvScripts.ItemDrag += new System.Windows.Forms.ItemDragEventHandler(this.tvScripts_ItemDrag);
			this.tvScripts.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tvScripts_AfterSelect);
			this.tvScripts.DragDrop += new System.Windows.Forms.DragEventHandler(this.tvScripts_DragDrop);
			this.tvScripts.DragEnter += new System.Windows.Forms.DragEventHandler(this.tvScripts_DragEnter);
			// 
			// imageList1
			// 
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			this.imageList1.Images.SetKeyName(0, "dots.png");
			this.imageList1.Images.SetKeyName(1, "blank.png");
			this.imageList1.Images.SetKeyName(2, "folder.png");
			// 
			// CodeEditorHost
			// 
			this.CodeEditorHost.Dock = System.Windows.Forms.DockStyle.Fill;
			this.CodeEditorHost.Location = new System.Drawing.Point(0, 0);
			this.CodeEditorHost.Name = "CodeEditorHost";
			this.CodeEditorHost.Size = new System.Drawing.Size(530, 419);
			this.CodeEditorHost.TabIndex = 0;
			this.CodeEditorHost.Text = "elementHost1";
			this.CodeEditorHost.Child = this.CodeEditor;
			// 
			// btnJSV8
			// 
			this.btnJSV8.Name = "btnJSV8";
			this.btnJSV8.Size = new System.Drawing.Size(188, 30);
			this.btnJSV8.Text = "New JsV8 Script";
			// 
			// frmScript
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(800, 450);
			this.Controls.Add(this.splitContainer1);
			this.Controls.Add(this.toolStrip1);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "frmScript";
			this.Text = "Scripts";
			this.TopMost = true;
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Me_FormClosing);
			this.Load += new System.EventHandler(this.frmScripts_Load);
			this.toolStrip1.ResumeLayout(false);
			this.toolStrip1.PerformLayout();
			this.splitContainer1.Panel1.ResumeLayout(false);
			this.splitContainer1.Panel2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
			this.splitContainer1.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.ToolStrip toolStrip1;
		private System.Windows.Forms.SplitContainer splitContainer1;
		private System.Windows.Forms.TreeView tvScripts;
		private System.Windows.Forms.ToolStripButton btnRun;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
		private System.Windows.Forms.ToolStripButton btnDelete;
		private System.Windows.Forms.ToolStripDropDownButton btnNew;
		private System.Windows.Forms.ToolStripButton btnHelp;
		private System.Windows.Forms.ToolStripMenuItem btnJS;
		private System.Windows.Forms.ToolStripMenuItem btnVB;
		private System.Windows.Forms.Integration.ElementHost CodeEditorHost;
		private CodeEditor CodeEditor;
		private System.Windows.Forms.ToolStripLabel lblStatus;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.ToolStripMenuItem btnFolder;
		private System.Windows.Forms.ToolStripSplitButton btnSave;
		private System.Windows.Forms.ToolStripMenuItem btnSaveAs;
		private System.Windows.Forms.ToolStripMenuItem btnJSV8;
	}
}