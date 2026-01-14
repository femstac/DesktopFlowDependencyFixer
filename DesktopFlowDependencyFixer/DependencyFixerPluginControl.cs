using McTools.Xrm.Connection;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

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

		#region Plugin Metadata
		public new string Name => "Desktop Flow Dependency Fixer";
		public string Description => "A tool to diagnose and move missing components into a target unmanaged solution.";
		public string Author => "Oluwafemi Tosin Ajigbayi";
		public string Version => "1.0.4";
		public string HelpUrl => "https://github.com/femstac/DesktopFlowDependencyFixer";
		public string SmallImageBase64 => null;
		public string BigImageBase64 => null;
		public int ConnectionMasterId => 1;
		public int SecondaryConnectionId => -1;
		#endregion

		#region UI Controls
		private ToolStrip toolStripMenu;
		private ToolStripButton tsbClose;
		private ToolStripButton tsbLoadSolutions;
		private ToolStripButton tsbHelp;

		// Layout Controls
		private Panel pnlInputContainer;
		private Splitter splitter;
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

			this.pnlInputContainer = new Panel();
			this.splitter = new Splitter();

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
			this.gbInput.SuspendLayout();
			this.tlpInput.SuspendLayout();
			this.toolStripMenu.SuspendLayout();
			this.pnlInputContainer.SuspendLayout();

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

			this.tsbHelp.Text = "Help / How-To";
			this.tsbHelp.Click += (s, e) => System.Diagnostics.Process.Start(HelpUrl);

			// 2. Logger
			this.rtbLogger.Dock = DockStyle.Fill;
			this.rtbLogger.ReadOnly = true;
			this.rtbLogger.BackColor = System.Drawing.Color.White;
			this.rtbLogger.Font = new System.Drawing.Font("Consolas", 9F);
			this.rtbLogger.BorderStyle = BorderStyle.None;
			this.Controls.Add(this.rtbLogger);

			// 3. Splitter
			this.splitter.Dock = DockStyle.Top;
			this.splitter.Height = 5;
			this.splitter.BackColor = System.Drawing.Color.Silver;
			this.Controls.Add(this.splitter);

			// 4. Input Container
			this.pnlInputContainer.Dock = DockStyle.Top;
			this.pnlInputContainer.Height = 350;
			this.pnlInputContainer.Padding = new Padding(0, 0, 0, 5);
			this.Controls.Add(this.pnlInputContainer);

			// 5. GroupBox
			this.gbInput.Dock = DockStyle.Fill;
			this.gbInput.Text = "Component Details";
			this.gbInput.Padding = new Padding(10);
			this.pnlInputContainer.Controls.Add(this.gbInput);

			// 6. TableLayoutPanel
			this.tlpInput.Dock = DockStyle.Fill;
			this.tlpInput.ColumnCount = 2;
			this.tlpInput.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160F));
			this.tlpInput.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
			this.tlpInput.RowCount = 6;

			this.tlpInput.RowStyles.Add(new RowStyle(SizeType.Absolute, 80F));
			this.tlpInput.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
			this.tlpInput.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
			this.tlpInput.RowStyles.Add(new RowStyle(SizeType.Absolute, 35F));
			this.tlpInput.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F));
			this.tlpInput.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
			this.gbInput.Controls.Add(this.tlpInput);

			// 7. Controls
			this.lblInstructions.Text = "Instructions:\n1. Locate the missing component's GUID from the error log or Dataverse URL.\n2. Select the Component Type (usually 'Desktop Flow Binary').\n3. Select the target unmanaged solution to inject the component into.";
			this.lblInstructions.Dock = DockStyle.Fill;
			this.lblInstructions.AutoSize = true;
			this.lblInstructions.ForeColor = System.Drawing.Color.DimGray;
			this.tlpInput.SetColumnSpan(this.lblInstructions, 2);
			this.tlpInput.Controls.Add(this.lblInstructions, 0, 0);

			this.lblComponentId.Text = "Component ID (GUID):";
			this.lblComponentId.Dock = DockStyle.Left;
			this.lblComponentId.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.tlpInput.Controls.Add(this.lblComponentId, 0, 1);

			this.txtComponentId.Dock = DockStyle.Fill;
			this.txtComponentId.Anchor = AnchorStyles.Left | AnchorStyles.Right;
			this.toolTipInfo.SetToolTip(this.txtComponentId, "Paste the GUID (e.g., d5362aa0-...) here.");
			this.tlpInput.Controls.Add(this.txtComponentId, 1, 1);

			this.lblComponentType.Text = "Component Type:";
			this.lblComponentType.Dock = DockStyle.Left;
			this.lblComponentType.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.tlpInput.Controls.Add(this.lblComponentType, 0, 2);

			this.cmbComponentType.Dock = DockStyle.Fill;
			this.cmbComponentType.DropDownStyle = ComboBoxStyle.DropDownList;
			this.tlpInput.Controls.Add(this.cmbComponentType, 1, 2);

			this.lblTargetSolution.Text = "Target Solution:";
			this.lblTargetSolution.Dock = DockStyle.Left;
			this.lblTargetSolution.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.tlpInput.Controls.Add(this.lblTargetSolution, 0, 3);

			this.cmbSolutions.Dock = DockStyle.Fill;
			this.cmbSolutions.DropDownStyle = ComboBoxStyle.DropDownList;
			this.tlpInput.Controls.Add(this.cmbSolutions, 1, 3);

			this.btnMoveComponent.Text = "Move Component to Solution";
			this.btnMoveComponent.Size = new System.Drawing.Size(250, 30);
			this.btnMoveComponent.Anchor = AnchorStyles.Left;

			// CRITICAL: WIRING UP THE EVENT HANDLER
			this.btnMoveComponent.Click += BtnMoveComponent_Click;

			this.tlpInput.Controls.Add(this.btnMoveComponent, 1, 4);

			this.tlpInput.ResumeLayout(false);
			this.tlpInput.PerformLayout();
			this.gbInput.ResumeLayout(false);
			this.pnlInputContainer.ResumeLayout(false);
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
			try
			{
				Log("Initiating component move...");

				if (!Guid.TryParse(txtComponentId.Text, out Guid componentId))
				{
					MessageBox.Show("Invalid Component ID. Please enter a valid GUID.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Log("Error: Invalid GUID format.");
					return;
				}

				if (cmbComponentType.SelectedItem == null)
				{
					MessageBox.Show("Please select a Component Type.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Log("Error: No Component Type selected.");
					return;
				}
				int componentType = ((ComponentTypeOption)cmbComponentType.SelectedItem).Value;

				if (cmbSolutions.SelectedItem == null)
				{
					MessageBox.Show("Please select a target solution.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Log("Error: No Target Solution selected.");
					return;
				}

				var solutionUniqueName = ((SolutionOption)cmbSolutions.SelectedItem).UniqueName;

				Log($"Parameters validated. ID: {componentId}, Type: {componentType}, Solution: {solutionUniqueName}");

				WorkAsync(new WorkAsyncInfo
				{
					Message = $"Adding component (Type: {componentType}) to '{solutionUniqueName}'...",
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
						if (args.Error != null)
						{
							Log($"API ERROR: {args.Error.Message}");
							MessageBox.Show($"Error executing request: {args.Error.Message}", "API Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
							return;
						}

						var request = (AddSolutionComponentRequest)args.Result;
						Log($"SUCCESS: Component '{request.ComponentId}' was added to solution '{request.SolutionUniqueName}'.");
						MessageBox.Show("Component added successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
				});
			}
			catch (Exception ex)
			{
				Log($"CRITICAL ERROR: {ex.Message}");
				MessageBox.Show($"An unexpected error occurred: {ex.Message}", "Critical Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void Log(string message) { rtbLogger.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n"); rtbLogger.ScrollToCaret(); }
		public override void ClosingPlugin(PluginCloseInfo info) { }
	}
}