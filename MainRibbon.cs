using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
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
			return GetResourceText("VstoOutlookAddInTemplate.MainRibbon.xml");
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
			Assembly asm = Assembly.GetExecutingAssembly();
			string[] resourceNames = asm.GetManifestResourceNames();
			for (int i = 0; i < resourceNames.Length; ++i)
			{
				if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
				{
					using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
					{
						if (resourceReader != null)
						{
							return resourceReader.ReadToEnd();
						}
					}
				}
			}
			return null;
		}

		#endregion
	}
}
