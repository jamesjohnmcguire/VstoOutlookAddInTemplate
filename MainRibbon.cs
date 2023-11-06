using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

namespace VstoOutlookAddInTemplate
{
	[ComVisible(true)]
	public class MainRibbon : Office.IRibbonExtensibility
	{
		private Office.IRibbonUI ribbon;

		public MainRibbon()
		{
		}

		#region IRibbonExtensibility Members

		public string GetCustomUI(string ribbonID)
		{
			string text = null;

			switch (ribbonID)
			{
				case "Microsoft.Outlook.Explorer":
					text = GetResourceText(
						"VstoOutlookAddInTemplate.MainRibbon.xml");
					break;
				case "Microsoft.Outlook.Mail.Compose":
				default:
					break;
			}

			return text;
		}

		#endregion

		#region Ribbon Callbacks
		//Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

		public void Ribbon_Load(Office.IRibbonUI ribbonUI)
		{
			this.ribbon = ribbonUI;
		}

		#endregion

		#region Helpers

		private static string GetResourceText(string resourceName)
		{
			string text = null;

			Assembly thisAssembly = Assembly.GetExecutingAssembly();

			using (Stream templateObjectStream =
				thisAssembly.GetManifestResourceStream(resourceName))
			{
				using (StreamReader reader = new StreamReader(templateObjectStream))
				{
					text = reader.ReadToEnd();
				}
			}

			return text;
		}

	#endregion
}
}
