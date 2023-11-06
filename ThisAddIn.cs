/////////////////////////////////////////////////////////////////////////////
// <copyright file="ThisAddIn.cs" company="James John McGuire">
// Copyright © 2023 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace VstoOutlookAddInTemplate
{
	/// <summary>
	/// Main Class for Add In.
	/// </summary>
	/// <seealso cref="Microsoft.Office.Tools.Outlook.OutlookAddInBase" />
	public partial class ThisAddIn
	{
		protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
		{
			return new MainRibbon();
		}

		private void ThisAddIn_Startup(object sender, System.EventArgs e)
		{
		}

		private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
		{
			// Note: Outlook no longer raises this event. If you have code
			// that must run when Outlook shuts down, see
			// https://go.microsoft.com/fwlink/?LinkId=506785
		}

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(ThisAddIn_Startup);
			this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
		}
	}
}
