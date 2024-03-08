/*
****************************************************************************
*  Copyright (c) 2024,  Skyline Communications NV  All Rights Reserved.    *
****************************************************************************

By using this script, you expressly agree with the usage terms and
conditions set out below.
This script and all related materials are protected by copyrights and
other intellectual property rights that exclusively belong
to Skyline Communications.

A user license granted for this script is strictly for personal use only.
This script may not be used in any way by anyone without the prior
written consent of Skyline Communications. Any sublicensing of this
script is forbidden.

Any modifications to this script by the user are only allowed for
personal use and within the intended purpose of the script,
and will remain the sole responsibility of the user.
Skyline Communications will not be responsible for any damages or
malfunctions whatsoever of the script resulting from a modification
or adaptation by the user.

The content of this script is confidential information.
The user hereby agrees to keep this confidential information strictly
secret and confidential and not to disclose or reveal it, in whole
or in part, directly or indirectly to any person, entity, organization
or administration without the prior written consent of
Skyline Communications.

Any inquiries can be addressed to:

	Skyline Communications NV
	Ambachtenstraat 33
	B-8870 Izegem
	Belgium
	Tel.	: +32 51 31 35 69
	Fax.	: +32 51 31 01 29
	E-mail	: info@skyline.be
	Web		: www.skyline.be
	Contact	: Ben Vandenberghe

****************************************************************************
Revision History:

DATE		VERSION		AUTHOR			COMMENTS

08-03-2024	1.0.0.1		JSE, Skyline	Initial version
****************************************************************************
*/

namespace SLC_AS_IDP_CheckElementProperties_1
{
	using System;
	using System.Collections.Generic;
	using System.IO;
	using System.Linq;
	using System.Text;

	using Newtonsoft.Json;

	using Skyline.DataMiner.Automation;
	using Skyline.DataMiner.Core.DataMinerSystem.Automation;
	using Skyline.DataMiner.Core.DataMinerSystem.Common;
	using Skyline.DataMiner.Core.DataMinerSystem.Common.Properties;
	using Skyline.DataMiner.Net.Messages;

	/// <summary>
	/// Represents a DataMiner Automation script.
	/// </summary>
	public class Script
	{
		private const string folder = @"C:\Skyline_Data\IDP Investigation\";

		private const bool logAll = false;
		private const bool ignorePropertySetToFalseAndUnmanaged = true;
		private const bool listToFix = true;
		private const bool fix = true;
		private const bool makeManageAgain = true;

		private const string viewToGetElements = "";

		/// <summary>
		/// The script entry point.
		/// </summary>
		/// <param name="engine">Link with SLAutomation process.</param>
		public void Run(IEngine engine)
		{
			engine.SetFlag(RunTimeFlags.NoKeyCaching);
			string fileName = DateTime.UtcNow.ToString("s").Replace(':', '_');
			string filePath = Path.Combine(folder, fileName + ".txt");
			string listPath = Path.Combine(folder, fileName + "ListToFix.csv");

			IDms thisDms = engine.GetDms();
			var idp = thisDms.GetElement("DataMiner IDP");
			Dictionary<string, string> managedElements = GetEntries(idp, 1100, 1104);
			Dictionary<string, string> unmanagedElements = GetEntries(idp, 1900, 1903);
			var elements = GetElements(engine, thisDms);
			StringBuilder sb = new StringBuilder();
			StringBuilder sbToList = new StringBuilder();

			List<string> toManage = new List<string>();

			foreach (var element in elements)
			{
				string id = element.DmsElementId.Value;
				string name = element.Name;
				IDmsElementProperty dmsElementProperty = element.Properties["IDP"];
				string idpPropertyValue = dmsElementProperty.Value;
				bool propertyThinksIsManaged = idpPropertyValue?.Equals("true", StringComparison.OrdinalIgnoreCase) ?? false;
				bool isManaged = true;
				if (!managedElements.TryGetValue(id, out string nameInIDP))
				{
					isManaged = false;
					nameInIDP = "<not managed by IDP>";
				}

				bool isUnmanaged = true;
				if (!unmanagedElements.TryGetValue(id, out string unnameInIDP))
				{
					isUnmanaged = false;
					unnameInIDP = "<not unmanaged by IDP>";
				}

				if (logAll ||
					(
						(propertyThinksIsManaged && !isManaged)
						|| (!ignorePropertySetToFalseAndUnmanaged && !propertyThinksIsManaged && !isUnmanaged)
						))
				{
					var dataBlock = new
					{
						Id = id,
						Name = name,
						NameInIDP = nameInIDP,
						NameInUnmanaged = unnameInIDP,
						IDPProperty = idpPropertyValue,
					};

					string dataString = JsonConvert.SerializeObject(dataBlock);
					sb.AppendLine(dataString);
				}

				if (propertyThinksIsManaged && !isUnmanaged && !isManaged)
				{
					if (fix)
					{
						var writeProp = dmsElementProperty as IWritableProperty;
						writeProp.Value = String.Empty;
						element.Update();
						engine.GenerateInformation($"Cleaning up {name}");
						toManage.Add(id);
					}

					if (listToFix)
					{
						sbToList.AppendLine($"{id},{name}");
					}
				}
			}


			var refreshButton = idp.GetStandaloneParameter<string>(72);
			refreshButton.SetValue("1");

			if (!Directory.Exists(folder))
			{
				Directory.CreateDirectory(folder);
			}

			File.WriteAllText(filePath, sb.ToString());

			if (sbToList.Length > 0)
			{
				File.WriteAllText(listPath, sbToList.ToString());
			}

			if (makeManageAgain && toManage.Any())
			{
				engine.Sleep(1000);
				var manageButton = idp.GetStandaloneParameter<string>(14);
				string toManageString = String.Join("|", toManage);
				engine.GenerateInformation($"Setting {toManageString}");
				manageButton.SetValue(toManageString);
			}
		}

		private static IEnumerable<IDmsElement> GetElements(IEngine engine, IDms dms)
		{
			var elements = dms.GetElements();
			if (String.IsNullOrWhiteSpace(viewToGetElements))
			{
				return elements;
			}

			var elementView = dms.GetView(viewToGetElements);

			var request = new GetLiteElementInfo()
			{
				ViewID = elementView.Id,
				ExcludeSubViews = false,
				IncludePaused = true,
				IncludeStopped = true,
			};

			var response = engine.SendSLNetMessage(request).OfType<LiteElementInfoEvent>().ToList();
			return elements.Where(element =>
				response.Any(innerElement =>
					element.AgentId == innerElement.DataMinerID
					&& element.Id == innerElement.ElementID));
		}

		private static Dictionary<string, string> GetEntries(IDmsElement idp, int tableId, int columnId)
		{
			IDmsTable managedTable = idp.GetTable(tableId);
			var namesColumn = managedTable.GetColumn<string>(columnId);
			var managedIds = managedTable.GetPrimaryKeys();
			var managedElements = managedIds.ToDictionary(id => id, id => namesColumn.GetValue(id, KeyType.PrimaryKey));
			return managedElements;
		}
	}
}