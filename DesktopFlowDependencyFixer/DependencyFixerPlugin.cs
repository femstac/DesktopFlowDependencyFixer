using DesktopFlowDependencyFixer;
using McTools.Xrm.Connection;
using Microsoft.Xrm.Sdk;
using System.ComponentModel.Composition;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace DesktopFlowDependencyFixer
{
	[Export(typeof(IXrmToolBoxPlugin))]
	[ExportMetadata("Name", "Desktop Flow Dependency Fixer")]
	[ExportMetadata("Description", "Diagnose and fix missing 'desktopflowbinary' dependencies by moving them to specific solutions.")]
	[ExportMetadata("SmallImageBase64", "")]
	[ExportMetadata("BigImageBase64", "")]
	[ExportMetadata("BackgroundColor", "White")]
	[ExportMetadata("PrimaryFontColor", "Black")]
	[ExportMetadata("SecondaryFontColor", "Gray")]
	public class DependencyFixerPlugin : PluginBase
	{
		public override IXrmToolBoxPluginControl GetControl()
		{
			return new DependencyFixerPluginControl();
		}
	}
}