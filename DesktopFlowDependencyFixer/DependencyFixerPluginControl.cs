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
		public string Version => "1.4.5";
		public string HelpUrl => "https://github.com/femstac/DesktopFlowDependencyFixer.git";
		public string SmallImageBase64 => null;
		public string BigImageBase64 => null;
		public int ConnectionMasterId => 1;
		public int SecondaryConnectionId => -1;
		#endregion

		#region UI Controls
		private ToolStrip toolStripMenu;
		private ToolStripButton tsbClose;
		private ToolStripButton tsbLoadSolutions;


		private SplitContainer splitContainer;
		private GroupBox gbInput;
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


			this.splitContainer = new SplitContainer();

			this.gbInput = new GroupBox();

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


			this.toolStripMenu.SuspendLayout();
			this.toolStripMenu.Items.AddRange(new ToolStripItem[] { this.tsbClose, this.tsbLoadSolutions });
			this.toolStripMenu.Location = new System.Drawing.Point(0, 0);
			this.toolStripMenu.Name = "toolStripMenu";
			this.toolStripMenu.Size = new System.Drawing.Size(800, 25);
			this.toolStripMenu.TabIndex = 0;
			this.Controls.Add(this.toolStripMenu);

			this.tsbClose.Text = "Close";
			this.tsbClose.Click += (sender, e) => CloseTool();

			this.tsbLoadSolutions.Text = "Load Unmanaged Solutions";
			this.tsbLoadSolutions.Click += TsbLoadSolutions_Click;


			this.splitContainer.Dock = DockStyle.Fill;
			this.splitContainer.Location = new System.Drawing.Point(0, 25);
			this.splitContainer.Name = "splitContainer";
			this.splitContainer.Orientation = Orientation.Horizontal;
			this.splitContainer.Size = new System.Drawing.Size(800, 575);
			this.splitContainer.SplitterDistance = 220;
			this.splitContainer.TabIndex = 1;
			this.Controls.Add(this.splitContainer);


			this.gbInput.SuspendLayout();
			this.gbInput.Dock = DockStyle.Fill;
			this.gbInput.Text = "Component Details";
			this.gbInput.Padding = new Padding(10);
			this.splitContainer.Panel1.Controls.Add(this.gbInput);


			this.lblComponentId.Text = "Component ID (GUID)";
			this.lblComponentId.Location = new System.Drawing.Point(10, 30);
			this.lblComponentId.Size = new System.Drawing.Size(150, 20);
			this.lblComponentId.AutoSize = false;

			this.txtComponentId.Location = new System.Drawing.Point(170, 30);
			this.txtComponentId.Size = new System.Drawing.Size(400, 20);
			this.txtComponentId.Anchor = AnchorStyles.Top | AnchorStyles.Left;

			this.gbInput.Controls.Add(this.lblComponentId);
			this.gbInput.Controls.Add(this.txtComponentId);


			this.lblComponentType.Text = "Component Type";
			this.lblComponentType.Location = new System.Drawing.Point(10, 60);
			this.lblComponentType.Size = new System.Drawing.Size(150, 20);
			this.lblComponentType.AutoSize = false;

			this.cmbComponentType.Location = new System.Drawing.Point(170, 60);
			this.cmbComponentType.Size = new System.Drawing.Size(400, 21);
			this.cmbComponentType.DropDownStyle = ComboBoxStyle.DropDownList;
			this.cmbComponentType.Anchor = AnchorStyles.Top | AnchorStyles.Left;

			this.gbInput.Controls.Add(this.lblComponentType);
			this.gbInput.Controls.Add(this.cmbComponentType);


			this.lblTargetSolution.Text = "Target Unmanaged Solution";
			this.lblTargetSolution.Location = new System.Drawing.Point(10, 90);
			this.lblTargetSolution.Size = new System.Drawing.Size(150, 20);
			this.lblTargetSolution.AutoSize = false;

			this.cmbSolutions.Location = new System.Drawing.Point(170, 90);
			this.cmbSolutions.Size = new System.Drawing.Size(400, 21);
			this.cmbSolutions.DropDownStyle = ComboBoxStyle.DropDownList;
			this.cmbSolutions.Anchor = AnchorStyles.Top | AnchorStyles.Left;

			this.gbInput.Controls.Add(this.lblTargetSolution);
			this.gbInput.Controls.Add(this.cmbSolutions);


			this.btnMoveComponent.Text = "Move Component to Solution";
			this.btnMoveComponent.Location = new System.Drawing.Point(170, 130);
			this.btnMoveComponent.Size = new System.Drawing.Size(250, 30);
			this.btnMoveComponent.Click += BtnMoveComponent_Click;
			this.gbInput.Controls.Add(this.btnMoveComponent);


			this.rtbLogger.Dock = DockStyle.Fill;
			this.rtbLogger.ReadOnly = true;
			this.rtbLogger.BackColor = System.Drawing.Color.White;
			this.rtbLogger.Font = new System.Drawing.Font("Consolas", 9F);
			this.splitContainer.Panel2.Controls.Add(this.rtbLogger);


			this.gbInput.ResumeLayout(false);
			this.splitContainer.Panel1.ResumeLayout(false);
			this.splitContainer.Panel2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.splitContainer)).EndInit();
			this.splitContainer.ResumeLayout(false);
			this.toolStripMenu.ResumeLayout(false);
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
						Log($"WARNING: Failed to load component definitions. Defaulting to standard values. Error: {args.Error.Message}");

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

					if (cmbComponentType.Items.Count > 0)
					{
						cmbComponentType.SelectedIndex = 0;
					}
					else
					{
						PopulateDropdownFallback();
					}
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
							Conditions =
						{
							new ConditionExpression("ismanaged", ConditionOperator.Equal, false),
							new ConditionExpression("isvisible", ConditionOperator.Equal, true)
						}
						}
					};
					args.Result = Service.RetrieveMultiple(query);
				},
				PostWorkCallBack = (args) =>
				{
					if (args.Error != null)
					{
						MessageBox.Show(args.Error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
						return;
					}
					var solutions = (EntityCollection)args.Result;



					var options = solutions.Entities.Select(s => new SolutionOption
					{
						FriendlyName = s.GetAttributeValue<string>("friendlyname"),
						UniqueName = s.GetAttributeValue<string>("uniquename")
					}).ToList();


					options = options.OrderBy(o => o.FriendlyName).ToList();

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

			if (!Guid.TryParse(txtComponentId.Text, out Guid componentId))
			{
				MessageBox.Show("Invalid Component ID. Please enter a valid GUID.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}


			if (cmbComponentType.SelectedItem == null)
			{
				MessageBox.Show("Please select a Component Type.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			int componentType = ((ComponentTypeOption)cmbComponentType.SelectedItem).Value;


			if (cmbSolutions.SelectedItem == null)
			{
				MessageBox.Show("Please select a target solution.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}


			var solutionUniqueName = ((SolutionOption)cmbSolutions.SelectedItem).UniqueName;


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
						Log($"ERROR: {args.Error.Message}");
						MessageBox.Show(args.Error.Message, "API Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
						return;
					}

					var request = (AddSolutionComponentRequest)args.Result;
					Log($"SUCCESS: Component '{request.ComponentId}' was added to solution '{request.SolutionUniqueName}'.");
					MessageBox.Show("Component added successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			});
		}

		private void Log(string message)
		{
			rtbLogger.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
			rtbLogger.ScrollToCaret();
		}


		public override void ClosingPlugin(PluginCloseInfo info) { }
	}

}
