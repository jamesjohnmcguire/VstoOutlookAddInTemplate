/////////////////////////////////////////////////////////////////////////////
// <copyright file="MainRibbon.cs" company="James John McGuire">
// Copyright © 2023 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
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

		public void OnAboutButton(Office.IRibbonControl control)
		{
			string version = GetVersion();

			string message = "Version: " + version;
			MessageBox.Show(message, "This Add In ");
		}

		public void OnRibbonLoad(Office.IRibbonUI ribbonUI)
		{
			this.ribbon = ribbonUI;
		}

		private static FileVersionInfo GetAssemblyInformation()
		{
			FileVersionInfo fileVersionInfo = null;

			Assembly assembly = Assembly.GetExecutingAssembly();

			string location = assembly.Location;

			if (string.IsNullOrWhiteSpace(location))
			{
				// Single file apps have no assemblies.
				Process process = Process.GetCurrentProcess();
				location = process.MainModule.FileName;
			}

			if (!string.IsNullOrWhiteSpace(location))
			{
				fileVersionInfo = FileVersionInfo.GetVersionInfo(location);
			}

			return fileVersionInfo;
		}

		private static string GetResourceText(string resourceName)
		{
			string text = null;

			Assembly thisAssembly = Assembly.GetExecutingAssembly();

			using (Stream templateObjectStream =
				thisAssembly.GetManifestResourceStream(resourceName))
			{
				if (templateObjectStream != null)
				{
					using (StreamReader reader =
						new StreamReader(templateObjectStream))
					{
						text = reader.ReadToEnd();
					}
				}
			}

			return text;
		}

		private static string GetVersion()
		{
			FileVersionInfo fileVersionInfo = GetAssemblyInformation();

			string assemblyVersion = fileVersionInfo.FileVersion;

			return assemblyVersion;
		}
	}
}
