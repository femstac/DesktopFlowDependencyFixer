using McTools.Xrm.Connection;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Data;
using System.Data.Common;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

namespace DesktopFlowDependencyFixer
{
	public class DependencyFixerPluginControl : PluginControlBase, IXrmToolBoxPluginControl
	{
		public class ComponentTypeOption
		{
			public string Name { get; set; }
			public int Value { get; set; }
			public ComponentTypeOption(string name, int value) { Name = name; Value = value; }
			public override string ToString() { return $"{Name} ({Value})"; }
		}

		public class SolutionOption
		{
			public string FriendlyName { get; set; }
			public string UniqueName { get; set; }
			public override string ToString() { return FriendlyName; }
		}

		#region Plugin Metadata Properties
		public new string Name => "Desktop Flow Dependency Fixer";
		public string Description => "A tool to diagnose and move missing components into a target unmanaged solution.";
		public string Author => "Oluwafemi Tosin Ajigbayi";
		public string Version => "1.5.2";
		public string HelpUrl => "https://github.com/femstac/DesktopFlowDependencyFixer";
		public string SmallImageBase64 => "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAMAAAAoLQ9TAAAAk1BMVEVHcEwLbrsAa74DhdUijdIEbb4KfMoCeswakNkGbb0Ebr4Ac8YSi9cJZ7EDZrghdroAa701ktBEn9UMfMoBecwOhNAfktgLitksl9kMh9MWhc0Kd8QAfs/d6vQCcMEJfs00kNDy9/v2+vwwh8c/ks4cgMaaxOLR4u5wrdgQjdk/ntmpz+fn8fc3mthJpt6q0+xIpNywMAo0AAAAH3RSTlMAFPz8FKZx9XREdfymBnAG9AYGRf52RfUVp0R3+//+n/R92AAAAJZJREFUGNN1j9cOgzAQBLENLrQQapptML3n/78uInGEeODeZqS9uzWMkwnDI5uua+50oylFKHngv6N9hpb3i0isRZr1wzrOkkQ6T9E8qYLnzhNsfHXRsipVTE0lrMsm7mgYC97ypoJfYZhJnPOua0tog9+SiMi8rksRMH0Fx8QRAgbQ1wJgZgtoMx/szwLLAscynndW+wORYwoKu8gWyAAAAABJRU5ErkJggg==";
		public string BigImageBase64 => "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAMAAABEpIrGAAABF1BMVEVHcEwkcK5Ll8gGbbwQb7gvkdEQbrgpjc4CabsNcbsOfckkjtMJhtUMarc2k84NbboDbr8Lg9EEb8AekdghkdcRi9gGargDaroDZbcHeMghlNsXkdsGe8wEdMYYh88Pcb0djNInlNcwldYXf8YMcb8QjNkXi9QKhNEEfMsSjNcPfssSgs0niswId8UAbL8AbsAAcMIEhtcChNUDf8/0+PwIidn7/Pzy9fkAc8Xm8PcCfM4Ae80AeMoAdcgBabwNjdr3+/3t8/j9/v0TecFUmsvU5PDg7PQBecsekNYBgdMxltYdfsRgrt5ws90zi8bQ5/NXo9ZQlMbG3Ou51ehard5aqNk9ndeuy+AbhMski89rq9Y8mNJloc31bfpJAAAALnRSTlMABAWHJicVLu5Si1P6LRY72u3qlIjqYpK4ubrt7e6IZGRjPj2W85fb89u5uTvbWVSwsQAAAZhJREFUOMvNk9dSwlAQhokJvQsIdhQLtvRATiohJBB6L+r7P4cbdZREhmv3KjPfv7v/7tkEAv8tDq5yuauDPYJgQRAKwd0sFYbMeIJhEvFAAAun/uSenOav8w8My7L3F88X2SdfndSJruu3AsN0OqzDZTlOevTWCJ/quiAwbHc06nIQknwW9trPfXJnydffZAl4O+Yb5hrqd+b9AUKD/mItt6sVn8k85M9rzVqt0eCNsaIoMU+DYBz8O/1m0+V1Yyyq6kuU+G1SLiTA33JQ++LGa3cjivTh5Y/gjmHYzohH3xxZ9qxH09TNloBlN1bdBg4ChCxr1qOoLUE5m3C44fAV/K1WCNnvQ5dvtQCT95ykLHh+apozZLdIispECWx7kBgsZ2wY0/V6alsgIHHfHipVSe61Wi3TnEwmJklG0l6OxWC7qkjTIk2Rbmg45n2sM1lRVBWmh/FAommRpEcQOgZ87nKwT0Y0TSuFvD2I43M8jWfc9Ax8REqE/6RCSWgadctHwVIytPswiUPwT+y5auyoWDzC/t3f+AEtHUwVOv1q1QAAAABJRU5ErkJggg==";
		public int ConnectionMasterId => 1;
		public int SecondaryConnectionId => -1;
		#endregion

		#region UI Controls
		private ToolStrip toolStripMenu;
		private ToolStripButton tsbClose;
		private ToolStripButton tsbLoadSolutions;
        private ToolStripButton tsbHelp;
		private SplitContainer splitContainer;
		private GroupBox gbInput;
        private TableLayoutPanel tlpInput;
        
        private System.Windows.Forms.Label lblInstructions; 
        private ToolTip toolTipInfo;
		private System.Windows.Forms.Label lblComponentId;
		private TextBox txtComponentId;
		private System.Windows.Forms.Label lblComponentType;
		private ComboBox cmbComponentType;
		private System.Windows.Forms.Label lblTargetSolution;
		private ComboBox cmbSolutions;
		private Button btnMoveComponent;
		private RichTextBox rtbLogger;
		#endregion

		public DependencyFixerPluginControl() { InitializeComponent(); }

		public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName, object parameter)
		{
			base.UpdateConnection(newService, detail, actionName, parameter);
			if (detail != null)
			{
				tsbLoadSolutions.Enabled = true;
				gbInput.Enabled = false;
				LoadEnvironmentComponentMetadata();
			}
		}
		private void InitializeComponent()
		{
			this.toolStripMenu = new ToolStrip();
			this.tsbClose = new ToolStripButton();
			this.tsbLoadSolutions = new ToolStripButton();
            this.tsbHelp = new ToolStripButton();
			this.splitContainer = new SplitContainer();
			this.gbInput = new GroupBox();
            this.tlpInput = new TableLayoutPanel();
            this.lblInstructions = new System.Windows.Forms.Label();
            this.toolTipInfo = new ToolTip();

			this.lblComponentId = new System.Windows.Forms.Label();
			this.txtComponentId = new TextBox();
			this.lblComponentType = new System.Windows.Forms.Label();
			this.cmbComponentType = new ComboBox();
			this.lblTargetSolution = new System.Windows.Forms.Label();
			this.cmbSolutions = new ComboBox();
			this.btnMoveComponent = new Button();
			this.rtbLogger = new RichTextBox();

			this.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.splitContainer)).BeginInit();
			this.splitContainer.Panel1.SuspendLayout();
			this.splitContainer.Panel2.SuspendLayout();
			this.splitContainer.SuspendLayout();
            this.gbInput.SuspendLayout();
            this.tlpInput.SuspendLayout();
			this.toolStripMenu.SuspendLayout();

            // 1. ToolStrip Setup
            this.toolStripMenu.Items.AddRange(new ToolStripItem[] { this.tsbClose, this.tsbLoadSolutions, new ToolStripSeparator(), this.tsbHelp });
			this.toolStripMenu.Location = new System.Drawing.Point(0, 0);
			this.toolStripMenu.Name = "toolStripMenu";
			this.toolStripMenu.Size = new System.Drawing.Size(800, 25);
			this.Controls.Add(this.toolStripMenu);

			this.tsbClose.Text = "Close";
			this.tsbClose.Click += (sender, e) => CloseTool();

			this.tsbLoadSolutions.Text = "Load Unmanaged Solutions";
			this.tsbLoadSolutions.Click += TsbLoadSolutions_Click;

            // HELP BUTTON CONFIG
            this.tsbHelp.Text = "Help / How-To";
            this.tsbHelp.Image = null;
            this.tsbHelp.Click += (s, e) => System.Diagnostics.Process.Start(HelpUrl);

            // 2. SplitContainer
            this.splitContainer.Dock = DockStyle.Fill;
            this.splitContainer.FixedPanel = FixedPanel.Panel1; 
            this.splitContainer.Orientation = Orientation.Horizontal;
            this.splitContainer.SplitterDistance = 280;
            this.Controls.Add(this.splitContainer);

            // 3. GroupBox
            this.gbInput.Dock = DockStyle.Fill;
            this.gbInput.Text = "Component Details";
            this.gbInput.Padding = new Padding(10);
            this.splitContainer.Panel1.Controls.Add(this.gbInput);

            // 4. TableLayoutPanel
            this.tlpInput.Dock = DockStyle.Fill;
            this.tlpInput.ColumnCount = 2;
            this.tlpInput.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160F));
            this.tlpInput.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            this.tlpInput.RowCount = 6;
            
            // Row Styles
            this.tlpInput.RowStyles.Add(new RowStyle(SizeType.Absolute, 60F)); // Instructions 
            this.tlpInput.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F)); // ID
            this.tlpInput.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F)); // Type
            this.tlpInput.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F)); // Solution
            this.tlpInput.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F)); // Button
            this.tlpInput.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); 
            this.gbInput.Controls.Add(this.tlpInput);

            // 5. Controls
            
            // Row 0: Instructions
            this.lblInstructions.Text = "Instructions:\n1. Locate the missing component's GUID from the error log or Dataverse URL.\n2. Select the Component Type (usually 'Desktop Flow Binary').\n3. Select the target unmanaged solution to inject the component into.";
            this.lblInstructions.Dock = DockStyle.Fill;
            this.lblInstructions.AutoSize = true;
            this.lblInstructions.ForeColor = System.Drawing.Color.DimGray;
            this.tlpInput.SetColumnSpan(this.lblInstructions, 2);
            this.tlpInput.Controls.Add(this.lblInstructions, 0, 0);

            // Row 1: ID
            this.lblComponentId.Text = "Component ID (GUID):";
            this.lblComponentId.Dock = DockStyle.Left;
            this.lblComponentId.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.tlpInput.Controls.Add(this.lblComponentId, 0, 1);

            this.txtComponentId.Dock = DockStyle.Fill; 
            this.txtComponentId.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            this.toolTipInfo.SetToolTip(this.txtComponentId, "Paste the GUID (e.g., d5362aa0-...) here.");
            this.tlpInput.Controls.Add(this.txtComponentId, 1, 1);

            // Row 2: Type
            this.lblComponentType.Text = "Component Type:";
            this.lblComponentType.Dock = DockStyle.Left;
            this.lblComponentType.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.tlpInput.Controls.Add(this.lblComponentType, 0, 2);

            this.cmbComponentType.Dock = DockStyle.Fill;
            this.cmbComponentType.DropDownStyle = ComboBoxStyle.DropDownList;
            this.tlpInput.Controls.Add(this.cmbComponentType, 1, 2);

            // Row 3: Solution
            this.lblTargetSolution.Text = "Target Solution:";
            this.lblTargetSolution.Dock = DockStyle.Left;
            this.lblTargetSolution.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.tlpInput.Controls.Add(this.lblTargetSolution, 0, 3);

            this.cmbSolutions.Dock = DockStyle.Fill;
            this.cmbSolutions.DropDownStyle = ComboBoxStyle.DropDownList;
            this.tlpInput.Controls.Add(this.cmbSolutions, 1, 3);

            // Row 4: Button
            this.btnMoveComponent.Text = "Move Component to Solution";
            this.btnMoveComponent.Size = new System.Drawing.Size(250, 30);
            this.btnMoveComponent.Anchor = AnchorStyles.Left;
            this.tlpInput.Controls.Add(this.btnMoveComponent, 1, 4);

            // Logger
            this.rtbLogger.Dock = DockStyle.Fill;
            this.rtbLogger.ReadOnly = true;
            this.rtbLogger.BackColor = System.Drawing.Color.White;
            this.rtbLogger.Font = new System.Drawing.Font("Consolas", 9F);
            this.rtbLogger.BorderStyle = BorderStyle.None;
            this.splitContainer.Panel2.Controls.Add(this.rtbLogger);

            this.tlpInput.ResumeLayout(false);
            this.tlpInput.PerformLayout();
            this.gbInput.ResumeLayout(false);
            this.splitContainer.Panel1.ResumeLayout(false);
            this.splitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer)).EndInit();
            this.splitContainer.ResumeLayout(false);
            this.toolStripMenu.ResumeLayout(false);
            this.toolStripMenu.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void LoadEnvironmentComponentMetadata()
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Retrieving environment component definitions...",
                Work = (worker, args) =>
                {
                    var query = new QueryExpression("solutioncomponentdefinition")
                    {
                        ColumnSet = new ColumnSet("solutioncomponenttype", "name", "primaryentityname"),
                        Criteria = new FilterExpression(LogicalOperator.Or)
                    };
                    query.Criteria.AddCondition("name", ConditionOperator.Equal, "desktopflowbinary");
                    query.Criteria.AddCondition("name", ConditionOperator.Equal, "workflow");
                    args.Result = Service.RetrieveMultiple(query);
                },
                PostWorkCallBack = (args) =>
                {
                    if (args.Error != null)
                    {
                        Log($"WARNING: Metadata load failed. Defaulting values. Error: {args.Error.Message}");
                        PopulateDropdownFallback();
                        return;
                    }
                    var results = (EntityCollection)args.Result;
                    cmbComponentType.Items.Clear();
                    foreach (var entity in results.Entities)
                    {
                        string name = entity.GetAttributeValue<string>("name");
                        int typeCode = entity.GetAttributeValue<int>("solutioncomponenttype");
                        string displayName = name == "desktopflowbinary" ? "Desktop Flow Binary" : "Process / Desktop Flow";
                        cmbComponentType.Items.Add(new ComponentTypeOption(displayName, typeCode));
                        Log($"Loaded Definition: {displayName} = {typeCode}");
                    }
                    if (cmbComponentType.Items.Count > 0) cmbComponentType.SelectedIndex = 0;
                    else PopulateDropdownFallback();
                }
            });
        }

        private void PopulateDropdownFallback()
        {
            cmbComponentType.Items.Clear();
            cmbComponentType.Items.Add(new ComponentTypeOption("Desktop Flow Binary (Fallback 10086)", 10086));
            cmbComponentType.Items.Add(new ComponentTypeOption("Process (Fallback 47)", 47));
            cmbComponentType.SelectedIndex = 0;
        }

        private void TsbLoadSolutions_Click(object sender, EventArgs e)
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Loading unmanaged solutions...",
                Work = (worker, args) =>
                {
                    var query = new QueryExpression("solution")
                    {
                        ColumnSet = new ColumnSet("uniquename", "friendlyname"),
                        Criteria = new FilterExpression
                        {
                            Conditions = { new ConditionExpression("ismanaged", ConditionOperator.Equal, false), new ConditionExpression("isvisible", ConditionOperator.Equal, true) }
                        }
                    };
                    args.Result = Service.RetrieveMultiple(query);
                },
                PostWorkCallBack = (args) =>
                {
                    if (args.Error != null) { MessageBox.Show(args.Error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                    var solutions = (EntityCollection)args.Result;
                    var options = solutions.Entities.Select(s => new SolutionOption
                    {
                        FriendlyName = s.GetAttributeValue<string>("friendlyname"),
                        UniqueName = s.GetAttributeValue<string>("uniquename")
                    }).OrderBy(o => o.FriendlyName).ToList();

                    cmbSolutions.DataSource = options;
                    cmbSolutions.DisplayMember = "FriendlyName";
                    cmbSolutions.ValueMember = "UniqueName";
                    gbInput.Enabled = true;
                    Log($"Successfully loaded {options.Count} unmanaged solutions.");
                }
            });
        }

        private void BtnMoveComponent_Click(object sender, EventArgs e)
        {
            if (!Guid.TryParse(txtComponentId.Text, out Guid componentId)) { MessageBox.Show("Invalid GUID.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
            if (cmbComponentType.SelectedItem == null) { MessageBox.Show("Select Type.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
            int componentType = ((ComponentTypeOption)cmbComponentType.SelectedItem).Value;
            if (cmbSolutions.SelectedItem == null) { MessageBox.Show("Select Solution.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
            var solutionUniqueName = ((SolutionOption)cmbSolutions.SelectedItem).UniqueName;

            WorkAsync(new WorkAsyncInfo
            {
                Message = $"Adding component...",
                Work = (worker, args) =>
                {
                    var request = new AddSolutionComponentRequest
                    {
                        ComponentId = componentId,
                        ComponentType = componentType,
                        SolutionUniqueName = solutionUniqueName,
                        AddRequiredComponents = true
                    };
                    Service.Execute(request);
                    args.Result = request;
                },
                PostWorkCallBack = (args) =>
                {
                    if (args.Error != null) { Log($"ERROR: {args.Error.Message}"); MessageBox.Show(args.Error.Message, "API Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                    var request = (AddSolutionComponentRequest)args.Result;
                    Log($"SUCCESS: Component added to '{request.SolutionUniqueName}'.");
                    MessageBox.Show("Success!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            });
        }

        private void Log(string message) { rtbLogger.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n"); rtbLogger.ScrollToCaret(); }
        public override void ClosingPlugin(PluginCloseInfo info) { }
    }
}