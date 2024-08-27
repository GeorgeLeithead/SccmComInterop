// Ignore Spelling: interop, sccm, wql, nal, embeded, sci, sys, res, netbios, obj

namespace SccmComInterop;

using System.Globalization;
using Microsoft.ConfigurationManagement.ManagementProvider;
using Microsoft.ConfigurationManagement.ManagementProvider.WqlQueryEngine;

/// <summary>COM interoperability Library for SCCM 2012</summary>
public static partial class CmInterop
{
	/// <summary>The add direct SCCM collection member machine.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="collectionId">The collection id.</param>
	/// <param name="resourceId">The resource id.</param>
	public static void AddDirectSccmCollectionMemberMachine(WqlConnectionManager connection, string collectionId, int resourceId)
	{
		try
		{
			// Get the specific collection instance to remove a population rule from.
			IResultObject collectionToModify = connection.GetInstance($"SMS_Collection.CollectionID='{collectionId}'");
			collectionToModify.Get();

			// Get the specific UserGroup object
			IResultObject system = connection.GetInstance($"SMS_R_System.ResourceID='{resourceId}'");
			List<IResultObject> collectionRules = collectionToModify.GetArrayItems("CollectionRules");
			bool found = collectionRules.Exists(collectionRule => collectionRule["RuleName"].StringValue.Equals(system["Name"].StringValue));
			if (found)
			{
				return;
			}

			IResultObject tempCollectionRule = connection.CreateEmbeddedObjectInstance("SMS_CollectionRuleDirect");

			tempCollectionRule["RuleName"].StringValue = system["Name"].StringValue;
			tempCollectionRule["ResourceClassName"].StringValue = "SMS_R_System";
			tempCollectionRule["ResourceID"].IntegerValue = system["ResourceID"].IntegerValue;

			collectionRules.Add(tempCollectionRule);

			collectionToModify.SetArrayItems("CollectionRules", collectionRules);
			collectionToModify.Put();

			// Now re-fresh the collection
			////RefreshSccmCollection(connection, collectionId); // Removed as in production is causing problems.
		}
		catch (SmsException ex)
		{
			throw new SmsException($"{ex.Message}, {ex.Details}");
		}
	}

	/// <summary>The add new computer.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="collectionId">The Collection ID</param>
	/// <param name="netBiosName">The net BIOS name.</param>
	/// <param name="systemManagementBiosGuid">The SM BIOS GUID.</param>
	/// <param name="macAddress">The mac address.</param>
	/// <returns>The <see cref="int"/>.</returns>
	/// <exception cref="ArgumentNullException">The System Management BIOS Name of MAC Address must be defined.</exception>
	public static int AddNewComputer(WqlConnectionManager connection, string collectionId, string netBiosName, string? systemManagementBiosGuid, string? macAddress)
	{
		if (systemManagementBiosGuid is null && macAddress is null)
		{
			throw new ArgumentNullException(netBiosName, "The System Management BIOS Name or MAC Address must be defined");
		}

		// Reformat macAddress to : separator.
		if (!string.IsNullOrEmpty(macAddress))
		{
			macAddress = macAddress.Replace("-", ":");
		}

		try
		{
			// Create the computer.
			Dictionary<string, object> inParams = new()
			{
															  { "NetbiosName", netBiosName },
															  { "SMBIOSGUID", systemManagementBiosGuid! },
															  { "MACAddress", macAddress! },
															  { "OverwriteExistingRecord", false }
														  };

			IResultObject outParams = connection.ExecuteMethod("SMS_Site", "ImportMachineEntry", inParams);

			// Add to All System collection.
			IResultObject collection = connection.GetInstance($"SMS_Collection.collectionId='{collectionId}'");
			IResultObject collectionRule = connection.CreateEmbeddedObjectInstance("SMS_CollectionRuleDirect");
			collectionRule["ResourceClassName"].StringValue = "SMS_R_System";
			collectionRule["ResourceID"].IntegerValue = outParams["ResourceID"].IntegerValue;
			Dictionary<string, object> inParams2 = new() { { "collectionRule", collectionRule } };
			_ = collection.ExecuteMethod("AddMembershipRule", inParams2);
			return outParams["ResourceID"].IntegerValue;
		}
		catch (SmsException ex)
		{
			throw new SmsException($"{ex.Message}, {ex.Details}");
		}
	}

	/// <summary>The add SCCM collection rule.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="collectionId">The collection id.</param>
	/// <param name="ruleName">The rule name.</param>
	/// <param name="wqlQuery">The WQL query.</param>
	/// <param name="limitToCollectionId">The limit to collection id.</param>
	public static void AddSccmCollectionRule(WqlConnectionManager connection, string collectionId, string ruleName, string wqlQuery, string limitToCollectionId)
	{
		try
		{
			// Get the specific collection instance to remove a population rule from.
			IResultObject collectionToModify = connection.GetInstance($"SMS_Collection.CollectionID='{collectionId}'");
			collectionToModify.Get();
			List<IResultObject> collectionRules = collectionToModify.GetArrayItems("CollectionRules");
			bool found = collectionRules.Exists(collectionRule => collectionRule["RuleName"].StringValue.Equals(ruleName));
			if (found)
			{
				return;
			}

			IResultObject tempCollectionRule = connection.CreateEmbeddedObjectInstance("SMS_CollectionRuleQuery");
			tempCollectionRule["RuleName"].StringValue = ruleName;
			tempCollectionRule["QueryExpression"].StringValue = wqlQuery;
			if (!string.IsNullOrWhiteSpace(limitToCollectionId))
			{
				tempCollectionRule["LimitToCollectionID"].StringValue = limitToCollectionId;
			}

			collectionRules.Add(tempCollectionRule);
			collectionToModify.SetArrayItems("CollectionRules", collectionRules);
			collectionToModify.Put();
		}
		catch (SmsException ex)
		{
			throw new SmsException($"{ex.Message}, {ex.Details}");
		}
	}

	/// <summary>The assign SCCM package to all distribution points.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="existingPackageId">The existing package id.</param>
	public static void AssignSccmPackageToAllDistributionPoints(WqlConnectionManager connection, string existingPackageId)
	{
		try
		{
			// Create the distribution point object (this is not an actual distribution point).
			IResultObject distributionPoint = connection.CreateInstance("SMS_DistributionPoint");

			// Associate the package with the new distribution point object.
			distributionPoint["PackageID"].StringValue = existingPackageId;

			// This query selects all distribution points
			const string Query = "SELECT * FROM SMS_SystemResourceList WHERE RoleName='SMS Distribution Point'";
			foreach (IResultObject resource in connection.QueryProcessor.ExecuteQuery(Query))
			{
				distributionPoint["ServerNALPath"].StringValue = resource["NALPath"].StringValue;
				distributionPoint["SiteCode"].StringValue = resource["SiteCode"].StringValue;

				// Save the distribution point object and properties.
				distributionPoint.Put();
				distributionPoint.Get();
			}
		}
		catch (SmsException ex)
		{
			throw new SmsException($"{ex.Message}, {ex.Details}");
		}
	}

	/// <summary>The assign SCCM package to distribution point.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="existingPackageId">The existing package id.</param>
	/// <param name="siteCode">The site code.</param>
	/// <param name="serverName">The server name.</param>
	/// <param name="nalPathQuery">The NAL path query.</param>
	public static void AssignSccmPackageToDistributionPoint(WqlConnectionManager connection, string existingPackageId, string siteCode, string serverName, string nalPathQuery)
	{
		try
		{
			// Create the distribution point object (this is not an actual distribution point).
			IResultObject distributionPoint = connection.CreateInstance("SMS_SCI_SysResUse");

			// Associate the package with the new distribution point object.
			distributionPoint["PackageID"].StringValue = existingPackageId;

			// This query selects a single distribution point based on the provided siteCode and serverName.
			string query = $"SELECT * FROM SMS_SystemResourceList WHERE RoleName='SMS Distribution Point' AND SiteCode='{siteCode}' AND ServerName='{serverName}' AND NALPAth Like '{nalPathQuery}'";
			foreach (IResultObject resource in connection.QueryProcessor.ExecuteQuery(query))
			{
				distributionPoint["ServerNALPath"].StringValue = resource["NALPath"].StringValue;
				distributionPoint["SiteCode"].StringValue = resource["SiteCode"].StringValue;
			}

			// Save the distribution point object and properties.
			distributionPoint.Put();
		}
		catch (SmsException ex)
		{
			throw new SmsException($"{ex.Message}, {ex.Details}");
		}
	}

	/// <summary>The assign SCCM package to distribution point group.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="existingPackageId">The existing package id.</param>
	/// <param name="distributionPointGroupName">The distribution point group name.</param>
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Minor Code Smell", "S3267:Loops should be simplified with \"LINQ\" expressions", Justification = "IResultObject' does not contain a definition for 'Cast' and the best extension method overload 'ParallelEnumerable.Cast<IResultObject>(ParallelQuery)' requires a receiver of type 'System.Linq.ParallelQuery'")]
	public static void AssignSccmPackageToDistributionPointGroup(WqlConnectionManager connection, string existingPackageId, string distributionPointGroupName)
	{
		try
		{
			// Create the distribution point object (this is not an actual distribution point).
			IResultObject distributionPoint = connection.CreateInstance("SMS_DistributionPoint");

			// Associate the package with the new distribution point object.
			distributionPoint["PackageID"].StringValue = existingPackageId;

			// This query selects a group of distribution points based on sGroupName
			string query = $"SELECT * FROM SMS_DistributionPointGroup WHERE sGroupName LIKE '{distributionPointGroupName}'";
			foreach (IResultObject resource in connection.QueryProcessor.ExecuteQuery(query))
			{
				foreach ((string nalPath, IResultObject listOfSystemResources) in from string nalPath in resource["arrNALPath"].StringArrayValue
																				  let SysResourceQuery = "SELECT * FROM SMS_SystemResourceList WHERE RoleName='SMS Distribution Point'"
																				  let listOfSystemResources = connection.QueryProcessor.ExecuteQuery(SysResourceQuery)
																				  select (nalPath, listOfSystemResources))
				{
					if (listOfSystemResources is null)
					{
						return;
					}

					foreach (IResultObject sysResource in listOfSystemResources)
					{
						if (sysResource["NALPath"].StringValue.Equals(nalPath))
						{
							distributionPoint["ServerNALPath"].StringValue = sysResource["NALPath"].StringValue;
							distributionPoint["SiteCode"].StringValue = sysResource["SiteCode"].StringValue;

							// Save the distribution point object and properties.
							distributionPoint.Put();
							distributionPoint.Get();
						}
					}
				}
			}
		}
		catch (SmsException ex)
		{
			throw new SmsException($"{ex.Message}, {ex.Details}");
		}
	}

	/// <summary>The connect SCCM server.</summary>
	/// <param name="serverName">The server name.</param>
	/// <param name="userName">The user name.</param>
	/// <param name="password">The password.</param>
	/// <returns>The <see cref="WqlConnectionManager"/>.</returns>
	public static WqlConnectionManager ConnectSccmServer(string serverName, string userName, string password)
	{
		WqlConnectionManager connection = new();
		_ = connection.Connect(serverName, userName, password);
		return connection;
	}

	/// <summary>The create SCCM collection.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="name">The name.</param>
	/// <param name="comment">The comment.</param>
	/// <param name="limitingCollectionId">The limiting collection ID</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject CreateSccmCollection(
		WqlConnectionManager connection,
		string name,
		string comment,
		string limitingCollectionId)
	{
		try
		{
			// Create a new SMS_Collection object.
			IResultObject collection = connection.CreateInstance("SMS_Collection");

			// Populate new collection properties.
			collection["Name"].StringValue = name;
			collection["Comment"].StringValue = comment;
			collection["OwnedByThisSite"].BooleanValue = true;
			collection["LimitToCollectionID"].StringValue = limitingCollectionId;

			// Save the new collection object and properties.
			// In this case, it seems necessary to 'get' the object again to access the properties
			collection.Put();
			collection.Get();

			return collection;
		}
		catch (SmsException ex)
		{
			throw new SmsException($"{ex.Message}, {ex.Details}");
		}
	}

	/// <summary>The create SCCM collection.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="name">The name.</param>
	/// <param name="comment">The comment.</param>
	/// <param name="recurWeeklyDay">Day of the week when the event is scheduled to occur. Sunday = 1, Monday = 2, etc.</param>
	/// <param name="recurWeeklyDayDuration">The number of days which the scheduled action occurs. Allowable values are in the range 0-31.  0 indicates that the scheduled action continues indefinitely.</param>
	/// <param name="recurWeeklyForNumberOfWeeks">Number of weeks for recurrence.  Allowable values are in the range 1-4.</param>
	/// <param name="recurWeeklyHourDuration">Number of hours during which the scheduled action occurs.  Allowable values are in the range 0-23. 0 indicating no duration.</param>
	/// <param name="recurWeeklyIsGmt">true if the time is in Coordinated Universal Time (UTC); otherwise local time.</param>
	/// <param name="recurWeeklyMinuteDuration">Number of minutes during which the scheduled action occurs.  Allowable values are in the range 0-59. 0 indicating no duration.</param>
	/// <param name="limitingCollectionId">The limiting collection ID</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject CreateSccmCollection(
		WqlConnectionManager connection,
		string name,
		string comment,
		int recurWeeklyDay,
		int recurWeeklyDayDuration,
		int recurWeeklyForNumberOfWeeks,
		int recurWeeklyHourDuration,
		bool recurWeeklyIsGmt,
		int recurWeeklyMinuteDuration,
		string limitingCollectionId)
	{
		try
		{
			// Create a new SMS_Collection object.
			IResultObject collection = connection.CreateInstance("SMS_Collection");

			// Populate new collection properties.
			collection["Name"].StringValue = name;
			collection["Comment"].StringValue = comment;
			collection["OwnedByThisSite"].BooleanValue = true;
			collection["LimitToCollectionID"].StringValue = limitingCollectionId;

			// Save the new collection object and properties.
			// In this case, it seems necessary to 'get' the object again to access the properties
			collection.Put();
			collection.Get();

			IResultObject recurWeekly = connection.CreateEmbeddedObjectInstance("SMS_ST_RecurWeekly");
			recurWeekly["Day"].IntegerValue = recurWeeklyDay;
			recurWeekly["DayDuration"].IntegerValue = recurWeeklyDayDuration;
			recurWeekly["ForNumberOfWeeks"].IntegerValue = recurWeeklyForNumberOfWeeks;
			recurWeekly["HourDuration"].IntegerValue = recurWeeklyHourDuration;
			recurWeekly["IsGMT"].BooleanValue = recurWeeklyIsGmt;
			recurWeekly["MinuteDuration"].IntegerValue = recurWeeklyMinuteDuration;
			recurWeekly["StartTime"].DateTimeValue = DateTime.Now;

			List<IResultObject> refreshSchedules = collection.GetArrayItems("RefreshSchedule");
			refreshSchedules.Add(recurWeekly);
			collection.SetArrayItems("RefreshSchedule", refreshSchedules);
			collection["RefreshType"].IntegerValue = 2; // Set the schedule type to 2 = Periodic (1 = Manual)
			collection.Put();
			collection.Get();

			return collection;
		}
		catch (SmsException ex)
		{
			throw new SmsException($"{ex.Message}, {ex.Details}");
		}
	}

	/// <summary>Create a SMS collection variable.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="name">The name.</param>
	/// <param name="value">The value.</param>
	/// <param name="mask">The mask.</param>
	/// <param name="collectionId">The collection id.</param>
	/// <param name="precedence">The precedence.</param>
	public static void CreateSccmCollectionVariable(WqlConnectionManager connection, string name, string value, bool mask, string collectionId, int precedence)
	{
		try
		{
			IResultObject? collectionSettings = null;

			// Get the collection settings. Create it if necessary.
			foreach (IResultObject setting in connection.QueryProcessor.ExecuteQuery($"Select * from SMS_CollectionSettings where CollectionID='{collectionId}'"))
			{
				setting.Get();
				collectionSettings = setting;
			}

			if (collectionSettings is null)
			{
				collectionSettings = connection.CreateInstance("SMS_CollectionSettings");
				collectionSettings["CollectionID"].StringValue = collectionId;
				collectionSettings.Put();
				collectionSettings.Get();
			}

			// Create the collection variable.
			List<IResultObject> collectionVariables = collectionSettings.GetArrayItems("CollectionVariables");
			IResultObject collectionVariable = connection.CreateEmbeddedObjectInstance("SMS_CollectionVariable");
			collectionVariable["Name"].StringValue = name;
			collectionVariable["Value"].StringValue = value;
			collectionVariable["IsMasked"].BooleanValue = mask;

			// Add the collection variable to the collection settings.
			collectionVariables.Add(collectionVariable);
			collectionSettings.SetArrayItems("CollectionVariables", collectionVariables);

			// Set the collection variable precedence.
			collectionSettings["CollectionVariablePrecedence"].IntegerValue = precedence;
			collectionSettings.Put();
		}
		catch (SmsException ex)
		{
			throw new SmsException($"{ex.Message}, {ex.Details}");
		}
	}

	/// <summary>Create a SMS device variable.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="resourceId">Device Resource Identifier.</param>
	/// <param name="name">The name.</param>
	/// <param name="value">The value.</param>
	/// <param name="mask">The mask.</param>
	public static void CreateSccmDeviceVariable(WqlConnectionManager connection, int resourceId, string name, string value, bool mask)
	{
		try
		{
			IResultObject? machineSettings = null;

			// Get the collection settings. Create it if necessary.
			foreach (IResultObject setting in connection.QueryProcessor.ExecuteQuery($"Select * from SMS_MachineSettings WHERE ResourceID='{resourceId}'"))
			{
				setting.Get();
				machineSettings = setting;
			}

			// Create the machine variables
			IResultObject machineVariable = connection.CreateEmbeddedObjectInstance("SMS_MachineVariable");
			machineVariable["Name"].StringValue = name;
			machineVariable["Value"].StringValue = value;
			machineVariable["IsMasked"].BooleanValue = mask;

			if (machineSettings is null)
			{
				machineSettings = connection.CreateInstance("SMS_MachineSettings");
				machineSettings["ResourceID"].StringValue = resourceId.ToString(CultureInfo.InvariantCulture);
				machineSettings["LocaleID"].IntegerValue = System.Threading.Thread.CurrentThread.CurrentUICulture.LCID;
				machineSettings["SourceSite"].StringValue = machineSettings.ConnectionManager.NamedValueDictionary["ConnectedSiteCode"].ToString();
			}

			List<IResultObject> machineVariables = machineSettings.GetArrayItems("MachineVariables");
			machineVariables.Add(machineVariable);
			machineSettings.SetArrayItems("MachineVariables", machineVariables);
			machineSettings.Put();
			machineSettings.Get();
		}
		catch (SmsException ex)
		{
			throw new SmsException($"{ex.Message}, {ex.Details}");
		}
	}

	/// <summary>Create a general-purpose embedded property used by the site control file to define the properties of a site control item.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="propertyName">Name of the property. The name is case sensitive and might contain several words, such as "Start-up Schedule".</param>
	/// <param name="value">A numeric value if the property is numeric. The default value is 0.</param>
	/// <param name="value1">A string value if the property is a string. The value is a registry data type if the property comes from the system registry. Otherwise, the value is the actual string for the property.</param>
	/// <param name="value2">A value to indicate the string value of the property if Value1 indicates a REG_SZ registry data type.</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject CreateSccmEmbededProperty(WqlConnectionManager connection, string propertyName, int value, string value1, string value2)
	{
		IResultObject ep = connection.CreateEmbeddedObjectInstance("SMS_EmbeddedProperty");
		ep["PropertyName"].StringValue = propertyName;
		ep["Value"].IntegerValue = value;
		ep["Value1"].StringValue = value1;
		ep["Value2"].StringValue = value2;
		return ep;
	}

	/// <summary>Create a general-purpose SCCM embedded object that defines property lists. The property lists are used by the site control file to define the string array properties of a site control item</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="propertyListName">Name of the property list. The name is case sensitive and might contain several words, for example, "Network Connection Accounts"</param>
	/// <param name="propertyValues">String values for the property list..</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject CreateSccmEmbededPropertyList(WqlConnectionManager connection, string propertyListName, string[] propertyValues)
	{
		IResultObject epl = connection.CreateEmbeddedObjectInstance("SMS_EmbeddedPropertyList");
		epl["PropertyListName"].StringValue = propertyListName;
		epl["Values"].StringArrayValue = propertyValues;
		return epl;
	}

	/// <summary>Create a specific usage of a server or other network resource</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="siteControlItemName">Unique name identifying the site control item.</param>
	/// <param name="siteCode">Unique three-letter site code identifying the site</param>
	/// <param name="roleName">Role of the server.</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject CreateSccmSciSysResUse(WqlConnectionManager connection, string siteControlItemName, string siteCode, string roleName)
	{
		try
		{
			IResultObject newSciSysReuseItem = connection.CreateInstance("SMS_SCI_SysResUse");

			// See http://msdn.microsoft.com/en-us/library/hh949718.aspx
			newSciSysReuseItem["ItemName"].StringValue = $"[\"Display=\\\\{siteControlItemName}\\\"]MSWNET:[\"SMS_SITE={siteCode}\"]\\\\{siteControlItemName}\\,SMS Site System";
			newSciSysReuseItem["ItemType"].StringValue = "System Resource Usage";
			newSciSysReuseItem["NALPath"].StringValue = $"[\"Display=\\\\{siteControlItemName}\\\"]MSWNET:[\"SMS_SITE={siteCode}\"]\\\\{siteControlItemName}\\";
			newSciSysReuseItem["NALType"].StringValue = "Windows NT Server";
			newSciSysReuseItem["NetworkOSPath"].StringValue = "\\\\" + siteControlItemName;
			newSciSysReuseItem["RoleName"].StringValue = roleName;
			newSciSysReuseItem["SiteCode"].StringValue = siteCode;

			// Save the server or other network resource
			newSciSysReuseItem.Put();
			newSciSysReuseItem.Get();

			return newSciSysReuseItem;
		}
		catch (SmsException ex)
		{
			throw new SmsException($"{ex.Message}, {ex.Details}");
		}
	}

	/// <summary>The get SCMM Boundary.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="boundaryName">The boundary name.</param>
	/// <param name="value">The values.</param>
	/// <param name="siteCode">The site Code.</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject GetSccmBoundary(WqlConnectionManager connection, string? boundaryName, string? value, string? siteCode)
	{
		string filter = $"DefaultSiteCode='{siteCode}' AND ";
		if (string.IsNullOrEmpty(siteCode))
		{
			filter = string.Empty;
		}

		if (!string.IsNullOrEmpty(boundaryName))
		{
			boundaryName = boundaryName.Replace("*", "%");
			string qualifier = boundaryName.Contains('%') ? " LIKE " : "=";
			filter = $"{filter} DisplayName{qualifier}'{boundaryName}'";
		}

		if (!string.IsNullOrEmpty(value))
		{
			value = value.Replace("*", "%");
			string qualifier = value.Contains('%') ? " LIKE " : "=";
			filter = $"{filter} Value{qualifier}'{value}'";
		}

		return GetSccmObject(connection, "SMS_Boundary", filter);
	}

	/// <summary>The get SCMM Boundary Group.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="boundaryName">The boundary Name.</param>
	/// <param name="siteCode">The site Code.</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject GetSccmBoundaryGroup(WqlConnectionManager connection, string? boundaryName, string siteCode)
	{
		string filter = $"DefaultSiteCode='{siteCode}'";
		if (!string.IsNullOrEmpty(boundaryName))
		{
			boundaryName = boundaryName.Replace("*", "%");
			string qualifier = boundaryName.Contains('%') ? " LIKE " : "=";
			filter = $"{filter} AND Name{qualifier}'{boundaryName}'";
		}

		return GetSccmObject(connection, "SMS_BoundaryGroup", filter);
	}

	/// <summary>The get SCMM Boundary Group from any site code.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="boundaryName">The boundary Name.</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject GetSccmBoundaryGroup(WqlConnectionManager connection, string boundaryName)
	{
		boundaryName = boundaryName.Replace("*", "%");
		string qualifier = boundaryName.Contains('%') ? " LIKE " : "=";
		string filter = $"Name{qualifier}'{boundaryName}'";
		return GetSccmObject(connection, "SMS_BoundaryGroup", filter);
	}

	/// <summary>The get SCCM  collection.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="filter">The filter.</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject GetSccmCollection(WqlConnectionManager connection, string filter) => GetSccmObject(connection, "SMS_Collection", filter);

	/// <summary>The get SCCM  collection members.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="collectionId">The collection id.</param>
	/// <param name="resourceName">The resource Name.</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject GetSccmCollectionMemberExists(WqlConnectionManager connection, string collectionId, string resourceName) => GetSccmObject(connection, "SMS_FullCollectionMembership", $"CollectionID = '{collectionId}' AND Name='{resourceName}'");

	/// <summary>The get SCCM  collection members.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="collectionId">The collection id.</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject GetSccmCollectionMembers(WqlConnectionManager connection, string collectionId) => GetSccmObject(connection, "SMS_FullCollectionMembership", $"CollectionID = '{collectionId}'");

	/// <summary>The get SCCM  computer.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="resourceId">The resource ID.</param>
	/// <param name="netbiosName">The NetBios name.</param>
	/// <param name="domainName">The domain name.</param>
	/// <param name="macAddress">MAC Address</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject GetSccmComputer(WqlConnectionManager connection, int? resourceId, string? netbiosName, string? domainName, string? macAddress)
	{
		string smsQueryFilter = string.Empty;
		if (resourceId is not null)
		{
			smsQueryFilter = $"ResourceID ='{resourceId}'";
		}

		if (!string.IsNullOrEmpty(netbiosName))
		{
			if (!string.IsNullOrWhiteSpace(smsQueryFilter))
			{
				smsQueryFilter += " AND ";
			}

			smsQueryFilter += "NetbiosName ";
			netbiosName = netbiosName.Replace("*", "%");
			if (netbiosName.Contains('%'))
			{
				smsQueryFilter += $"LIKE '{netbiosName}'";
			}
			else
			{
				smsQueryFilter += $"='{netbiosName}'";
			}
		}

		if (!string.IsNullOrEmpty(macAddress))
		{
			if (!string.IsNullOrWhiteSpace(smsQueryFilter))
			{
				smsQueryFilter += " AND ";
			}

			smsQueryFilter += $"MacAddresses ='{macAddress}'";
		}

		if (string.IsNullOrEmpty(domainName))
		{
			return GetSccmObject(connection, "SMS_R_System", smsQueryFilter);
		}

		if (!string.IsNullOrEmpty(domainName))
		{
			if (!string.IsNullOrWhiteSpace(smsQueryFilter))
			{
				smsQueryFilter += " AND ";
			}

			smsQueryFilter += "ResourceDomainOrWorkgroup ";
			domainName = domainName.Replace("*", "%");
			if (domainName.Contains('%'))
			{
				smsQueryFilter += $"LIKE '{domainName}'";
			}
			else
			{
				smsQueryFilter += $"='{domainName}'";
			}
		}

		return GetSccmObject(connection, "SMS_R_System", smsQueryFilter);
	}

	/// <summary>The get SCCM distribution point.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="siteSystemServerName">The site system server name.</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject GetSccmDistributionPoint(WqlConnectionManager connection, string siteSystemServerName)
	{
		siteSystemServerName = siteSystemServerName.Replace("*", "%");
		string qualifier = siteSystemServerName.Contains('%') ? " LIKE " : "=";
		string smsQueryFilter = $"RoleName='SMS Distribution Point' AND ServerRemoteName{qualifier}'{siteSystemServerName}'";
		return GetSccmObject(connection, "SMS_SystemResourceList", smsQueryFilter);
	}

	/// <summary>The get SCCM Distribution Point Group.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="distributionPointGroupName">The distribution point group name.</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject GetSccmDistributionPointGroup(WqlConnectionManager connection, string distributionPointGroupName)
	{
		distributionPointGroupName = distributionPointGroupName.Replace("*", "%");
		string qualifier = distributionPointGroupName.Contains('%') ? " LIKE " : "=";
		string filter = $"NAME{qualifier}'{distributionPointGroupName}'";
		return GetSccmObject(connection, "SMS_DistributionPointGroup", filter);
	}

	/// <summary>The get SCCM distribution point information.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="distributionPointGroupName">The distribution point group name.</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject GetSccmDistributionPointInfo(WqlConnectionManager connection, string distributionPointGroupName)
	{
		distributionPointGroupName = distributionPointGroupName.Replace("*", "%");
		string qualifier = distributionPointGroupName.Contains('%') ? " LIKE " : "=";
		string filter = $"NAME{qualifier}'{distributionPointGroupName}'";
		return GetSccmObject(connection, "SMS_DistributionPointInfo", filter);
	}

	/// <summary>Get a Distribution Point SCI Address.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="itemName">The item name.</param>
	/// <param name="siteCode">The site code.</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject GetSccmDistributionPointSciAddress(WqlConnectionManager connection, string itemName, string siteCode)
	{
		string addressInstanceQuery = $"SMS_SCI_ADDRESS.FileType=2,ItemName=\"{itemName}|MS_LAN\",ItemType=\"Address\",SiteCode=\"{siteCode}\"";
		return connection.GetInstance(addressInstanceQuery);
	}

	/// <summary>The get SCCM  object.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="objectClass">The object class.</param>
	/// <param name="filter">The filter.</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject GetSccmObject(WqlConnectionManager connection, string objectClass, string filter)
	{
		string query = filter?.Length == 0 ? $"SELECT * FROM {objectClass}" : $"SELECT * FROM {objectClass} WHERE {filter}";
		return connection.QueryProcessor.ExecuteQuery(query);
	}

	/// <summary>The get SCCM  package.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="filter">The filter.</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject GetSccmPackage(WqlConnectionManager connection, string filter) => GetSccmObject(connection, "SMS_Package", filter);

	/// <summary>The get SCCM  package distribution points.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="existingPackageId">The existing package id.</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject GetSccmPackageDistributionPoints(WqlConnectionManager connection, string existingPackageId) => GetSccmObject(connection, "SMS_DistributionPoint", $"PackageID='{existingPackageId}'");

	/// <summary>The get SCCM SCI Resource Use.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="serverName">The server name.</param>
	/// <param name="siteCode">The site code.</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject GetSccmSciResUse(WqlConnectionManager connection, string? serverName, string siteCode)
	{
		string filter = $"SiteCode='{siteCode}'";
		if (!string.IsNullOrEmpty(serverName))
		{
			serverName = serverName.Replace("*", "%");
			string qualifier = serverName.Contains('%') ? " LIKE " : "=";
			filter = $"{filter} AND ItemName{qualifier}'{serverName}'";
		}

		return GetSccmObject(connection, "SMS_SCI_SysResUse", filter);
	}

	/// <summary>The get SCCM SCI Resource Use.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="serverName">The server name.</param>
	/// <param name="siteCode">The site code.</param>
	/// <param name="roleName">The role name</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject GetSccmSciResUse(WqlConnectionManager connection, string? serverName, string siteCode, string? roleName)
	{
		string filter = $"SiteCode='{siteCode}'";
		if (!string.IsNullOrEmpty(serverName))
		{
			serverName = serverName.Replace("*", "%");
			string qualifier = serverName.Contains('%') ? " LIKE " : "=";
			filter = $"{filter} AND ItemName{qualifier}'{serverName}'";
		}

		if (!string.IsNullOrEmpty(roleName))
		{
			filter = $"{filter} AND RoleName='{roleName}'";
		}

		return GetSccmObject(connection, "SMS_SCI_SysResUse", filter);
	}

	/// <summary>Get all Configuration Manager site installations</summary>
	/// <param name="connection">The connection</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject GetSccmSite(WqlConnectionManager connection)
	{
		const string Query = "SELECT * FROM SMS_Site";
		return connection.QueryProcessor.ExecuteQuery(Query);
	}

	/// <summary>Get a collection SCCM site systems.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="siteSystemServerName">The site system server name.</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject GetSccmSiteSystem(WqlConnectionManager connection, string siteSystemServerName)
	{
		siteSystemServerName = siteSystemServerName.Replace("*", "%");
		string qualifier = siteSystemServerName.Contains('%') ? " LIKE " : "=";
		string smsQueryFilter = $"RoleName='SMS Site System' AND ServerRemoteName{qualifier}'{siteSystemServerName}'";
		return GetSccmObject(connection, "SMS_SystemResourceList", smsQueryFilter);
	}

	/// <summary>Get a collection SCCM site systems.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="siteSystemServerName">The site system server name.</param>
	/// <param name="siteCode">The site Code.</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject GetSccmSiteSystem(WqlConnectionManager connection, string siteSystemServerName, string? siteCode)
	{
		if (string.IsNullOrEmpty(siteCode))
		{
			return GetSccmSiteSystem(connection, siteSystemServerName);
		}

		siteSystemServerName = siteSystemServerName.Replace("*", "%");
		string qualifierSystemServerName = siteSystemServerName.Contains('%') ? " LIKE " : "=";
		siteCode = siteCode.Replace("*", "%");
		string qualifierSiteCode = siteCode.Contains('%') ? " LIKE " : "=";
		string smsQueryFilter = $"RoleName='SMS Site System' AND ServerRemoteName{qualifierSystemServerName}'{siteSystemServerName}' AND SiteCode{qualifierSiteCode}'{siteCode}'";
		return GetSccmObject(connection, "SMS_SystemResourceList", smsQueryFilter);
	}

	/// <summary>Get a collection of SCCM System Resource List items</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="siteSystemServerName">Server name.</param>
	/// <param name="siteCode">Site code.</param>
	/// <returns>The <see cref="IResultObject"/>.</returns>
	public static IResultObject GetSccmSystemResourceList(WqlConnectionManager connection, string siteSystemServerName, string? siteCode)
	{
		siteSystemServerName = siteSystemServerName.Replace("*", "%");
		string qualifier = siteSystemServerName.Contains('%') ? " LIKE " : "=";
		string smsQueryFilter = $"ServerRemoteName{qualifier}'{siteSystemServerName}'";

		if (!string.IsNullOrEmpty(siteCode))
		{
			string qualifierSiteCode = siteCode.Contains('%') ? " LIKE " : "=";
			smsQueryFilter += $" AND SiteCode{qualifierSiteCode}'{siteCode}'";
		}

		return GetSccmObject(connection, "SMS_SystemResourceList", smsQueryFilter);
	}

	/// <summary>The move console folder item.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="itemObjectId">The item object id.</param>
	/// <param name="objectType">The object type.</param>
	/// <param name="sourceContainerId">The source container id.</param>
	/// <param name="destinationContainerId">The destination container id.</param>
	public static void MoveConsoleFolderItem(WqlConnectionManager connection, string itemObjectId, FolderType objectType, int sourceContainerId, int destinationContainerId)
	{
		try
		{
			Dictionary<string, object> parameters = [];
			string[] sourceItems =
			[
				// Only one item is being moved, the array size is 1.
				itemObjectId,
			];
			parameters.Add("InstanceKeys", sourceItems);
			parameters.Add("ContainerNodeID", sourceContainerId);
			parameters.Add("TargetContainerNodeID", destinationContainerId);
			parameters.Add("ObjectType", (int)objectType);
			_ = connection.ExecuteMethod("SMS_ObjectContainerItem", "MoveMembers", parameters);
		}
		catch (SmsException ex)
		{
			throw new SmsException($"{ex.Message}, {ex.Details}");
		}
	}

	/// <summary>The refresh SCCM collection.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="collectionIdToRefresh">The collection id to refresh.</param>
	public static void RefreshSccmCollection(WqlConnectionManager connection, string collectionIdToRefresh)
	{
		try
		{
			// Get the specific collection instance to refresh.
			IResultObject collectionToRefresh = connection.GetInstance("SMS_Collection.CollectionID='" + collectionIdToRefresh + "'");

			// refresh the collection.
			_ = collectionToRefresh.ExecuteMethod("RequestRefresh", null);
		}
		catch (SmsException ex)
		{
			throw new SmsException($"{ex.Message}, {ex.Details}");
		}
	}

	/// <summary>The refresh SCCM package at all distribution points.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="existingPackageId">The existing package id.</param>
	public static void RefreshSccmPackageAtAllDistributionPoints(WqlConnectionManager connection, string existingPackageId)
	{
		try
		{
			// This query selects all SMS Distribution Point Servers
			const string Query = "SELECT * FROM SMS_SystemResourceList WHERE RoleName='SMS Distribution Point'";

			IResultObject listOfResources = connection.QueryProcessor.ExecuteQuery(Query);
			foreach (IResultObject resource in listOfResources)
			{
				string distributionPointQuery = $"SELECT * FROM SMS_DistributionPoint WHERE PackageID='{existingPackageId}'";
				IResultObject listOfDistributionPoints = connection.QueryProcessor.ExecuteQuery(distributionPointQuery);
				foreach (IResultObject dp in listOfDistributionPoints)
				{
					if (!dp["ServerNALPath"].StringValue.Equals(resource["NalPath"].StringValue))
					{
						continue;
					}

					// Set Refresh Flag
					dp["RefreshNow"].BooleanValue = true;
					dp.Put();
				}
			}
		}
		catch (SmsException ex)
		{
			throw new SmsException($"{ex.Message}, {ex.Details}");
		}
	}

	/// <summary>The refresh SCCM package at distribution point.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="existingPackageId">The existing package id.</param>
	/// <param name="siteCode">The site code.</param>
	/// <param name="serverName">The server name.</param>
	/// <param name="nalPathQuery">The NAL path query.</param>
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Minor Code Smell", "S3267:Loops should be simplified with \"LINQ\" expressions", Justification = "IResultObject' does not contain a definition for 'Cast' and the best extension method overload 'ParallelEnumerable.Cast<IResultObject>(ParallelQuery)' requires a receiver of type 'System.Linq.ParallelQuery'")]
	public static void RefreshSccmPackageAtDistributionPoint(WqlConnectionManager connection, string existingPackageId, string siteCode, string serverName, string nalPathQuery)
	{
		try
		{
			// This query selects System Resource
			string query = $"SELECT * FROM SMS_SystemResourceList WHERE RoleName='SMS Distribution Point' AND SiteCode='{siteCode}' AND ServerName='{serverName}' AND NALPath LIKE '{nalPathQuery}";

			IResultObject listOfResources = connection.QueryProcessor.ExecuteQuery(query);
			foreach (IResultObject resource in listOfResources)
			{
				string distributionPointQuery = $"SELECT * FROM SMS_DistributionPoint WHERE PackageID='{existingPackageId}'";
				IResultObject listOfDistributionPoints = connection.QueryProcessor.ExecuteQuery(distributionPointQuery);
				if (listOfDistributionPoints is null)
				{
					return;
				}

				foreach (IResultObject dp in listOfDistributionPoints)
				{
					if (dp["ServerNALPath"].StringValue.Equals(resource["NalPath"].StringValue))
					{
						// Set Refresh Flag
						dp["RefreshNow"].BooleanValue = true;
						dp.Put();
						break;
					}
				}
			}
		}
		catch (SmsException ex)
		{
			throw new SmsException($"{ex.Message}, {ex.Details}");
		}
	}

	/// <summary>The refresh SCCM package at distribution point group.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="existingPackageId">The existing package id.</param>
	/// <param name="distributionPointGroupName">The distribution point group name.</param>
	public static void RefreshSccmPackageAtDistributionPointGroup(WqlConnectionManager connection, string existingPackageId, string distributionPointGroupName)
	{
		try
		{
			// This query selects a group of distribution points based on sGroupName
			string query = $"SELECT * FROM SMS_DistributionPointGroup WHERE sGroupName LIKE '{distributionPointGroupName}'";

			foreach (IResultObject resource in connection.QueryProcessor.ExecuteQuery(query))
			{
				foreach (string nalPath in resource["arrNALPath"].StringArrayValue)
				{
					string distributionPointQuery = $"SELECT * FROM SMS_DistributionPoint WHERE PackageID='{existingPackageId}'";
					IResultObject listOfDistributionPoints = connection.QueryProcessor.ExecuteQuery(distributionPointQuery);
					foreach (IResultObject dp in listOfDistributionPoints)
					{
						if (!dp["ServerNALPath"].StringValue.Equals(nalPath))
						{
							continue;
						}

						// Set Refresh Flag
						dp["RefreshNow"].BooleanValue = true;
						dp.Put();
					}
				}
			}
		}
		catch (SmsException ex)
		{
			throw new SmsException($"{ex.Message}, {ex.Details}");
		}
	}

	/// <summary>The remove SCCM collection member.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="collectionId">The collection ID.</param>
	/// <param name="ruleName">The rule name.</param>
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Minor Code Smell", "S3267:Loops should be simplified with \"LINQ\" expressions", Justification = "IResultObject' does not contain a definition for 'Cast' and the best extension method overload 'ParallelEnumerable.Cast<IResultObject>(ParallelQuery)' requires a receiver of type 'System.Linq.ParallelQuery'")]
	public static void RemoveSccmCollectionMember(WqlConnectionManager connection, string collectionId, string ruleName)
	{
		try
		{
			// Get the specific collection instance to remove a population rule from.
			IResultObject collectionToModify = connection.GetInstance($"SMS_Collection.CollectionID='{collectionId}'");
			List<IResultObject>? collectionRules = collectionToModify.GetArrayItems("CollectionRules");
			if (collectionRules is null)
			{
				return;
			}

			foreach (IResultObject collectionRule in collectionRules)
			{
				if (collectionRule["RuleName"].StringValue.Equals(ruleName))
				{
					_ = collectionRules.Remove(collectionRule);

					collectionToModify.SetArrayItems("CollectionRules", collectionRules);
					collectionToModify.Put();

					break;
				}
			}
		}
		catch (SmsException ex)
		{
			throw new SmsException($"{ex.Message}, {ex.Details}");
		}
	}

	/// <summary>The remove SCCM object.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="filter">The filter.</param>
	/// <param name="objClass">The object class.</param>
	public static void RemoveSccmObject(WqlConnectionManager connection, string filter, string objClass)
	{
		foreach (IResultObject objToBeDeleted in GetSccmObject(connection, objClass, filter))
		{
			objToBeDeleted.Delete();
		}
	}

	/// <summary>The remove SCCM package from all distribution points.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="existingPackageId">The existing package ID.</param>
	public static void RemoveSccmPackageFromAllDistributionPoints(WqlConnectionManager connection, string existingPackageId)
	{
		try
		{
			// This query selects all distribution points
			string query = $"SELECT * FROM SMS_DistributionPoint WHERE PackageID='{existingPackageId}'";
			foreach (IResultObject resource in connection.QueryProcessor.ExecuteQuery(query))
			{
				// deletes the package DP
				resource.Delete();
			}
		}
		catch (SmsException ex)
		{
			throw new SmsException($"{ex.Message}, {ex.Details}");
		}
	}

	/// <summary>The remove SCCM package from distribution point.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="existingPackageId">The existing package ID.</param>
	/// <param name="siteCode">The site code.</param>
	/// <param name="serverName">The server name.</param>
	/// <param name="nalPathQuery">The NAL path query.</param>
	public static void RemoveSccmPackageFromDistributionPoint(WqlConnectionManager connection, string existingPackageId, string siteCode, string serverName, string nalPathQuery)
	{
		try
		{
			// This query selects the specified package from a single distribution point based on the provided siteCode and serverName.
			string query =
				$"SELECT * FROM SMS_SystemResourceList WHERE RoleName='SMS Distribution Point' AND SiteCode='{siteCode}' AND ServerName='{serverName}' AND PackageID='{existingPackageId}' AND NALPath LIKE '{nalPathQuery}'";
			foreach (IResultObject resource in connection.QueryProcessor.ExecuteQuery(query))
			{
				string distributionPointQuery = $"SELECT * FROM SMS_DistributionPoint WHERE PackageID='{existingPackageId}' AND ServerNALPath=\"{resource["NalPath"].StringValue}\"";
				foreach (IResultObject dp in connection.QueryProcessor.ExecuteQuery(distributionPointQuery))
				{
					// Delete Package Distribution Point
					dp.Delete();
				}
			}
		}
		catch (SmsException ex)
		{
			throw new SmsException($"{ex.Message}, {ex.Details}");
		}
	}

	/// <summary>The remove SCCM package from distribution point group.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="existingPackageId">The existing package ID.</param>
	/// <param name="distributionPointGroupName">The distribution point group name.</param>
	[System.Diagnostics.CodeAnalysis.SuppressMessage("Minor Code Smell", "S3267:Loops should be simplified with \"LINQ\" expressions", Justification = "IResultObject' does not contain a definition for 'Cast' and the best extension method overload 'ParallelEnumerable.Cast<IResultObject>(ParallelQuery)' requires a receiver of type 'System.Linq.ParallelQuery'")]
	public static void RemoveSccmPackageFromDistributionPointGroup(WqlConnectionManager connection, string existingPackageId, string distributionPointGroupName)
	{
		try
		{
			// This query selects a group of distribution points based on sGroupName
			string query = $"SELECT * FROM SMS_DistributionPointGroup WHERE sGroupName LIKE '{distributionPointGroupName}'";

			foreach (IResultObject resource in connection.QueryProcessor.ExecuteQuery(query))
			{
				foreach (string nalPath in resource["arrNALPath"].StringArrayValue)
				{
					string distributionPointResourceQuery = $"SELECT * FROM SMS_DistributionPoint WHERE PackageID='{existingPackageId}'";
					IResultObject? listOfDPs = connection.QueryProcessor.ExecuteQuery(distributionPointResourceQuery);
					if (listOfDPs is null)
					{
						return;
					}

					foreach (IResultObject distributionPoint in listOfDPs)
					{
						if (distributionPoint["ServerNALPath"].StringValue.Equals(nalPath))
						{
							// Delete the DP
							distributionPoint.Delete();
						}
					}
				}
			}
		}
		catch (SmsException ex)
		{
			throw new SmsException($"{ex.Message}, {ex.Details}");
		}
	}

	/// <summary>Run PXE clear on device.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="netBiosName">Device NetBIOSName</param>
	/// <exception cref="SmsException">Failed to clear PXE Advertisement.</exception>
	/// <exception cref="ArgumentNullException">No PXE Advertisement record for device.</exception>
	/// <returns>The <see cref="bool"/> result of clearing the PXE request.</returns>
	public static bool RunClearOnDevice(WqlConnectionManager connection, string netBiosName)
	{
		List<int> list = [];
		Dictionary<string, object> dictionary = [];
		string query = $"SELECT * FROM SMS_LastPXEAdvertisement WHERE NetBiosName='{netBiosName}'";
		IResultObject array = connection.QueryProcessor.ExecuteQuery(query);
		if (array is not null)
		{
			IResultObject computerObject = GetSccmComputer(connection, null, netBiosName, null, null) ?? throw new SmsException($"Unable to locate Resource ID for NetBiosName: NetBiosName: {netBiosName}");
			computerObject.Get();
			list.Add(computerObject["ResourceID"].IntegerValue);
		}

		if (list.Count == 0)
		{
			return true; // There was NO PXE Requests for the device, so technically it IS Clear!
		}

		dictionary.Add("ResourceIDs", list.ToArray());
		IResultObject resultObject = connection.ExecuteMethod("SMS_Collection", "ClearLastNBSAdvForMachines", dictionary) ?? throw new SmsException($"Failed to clear PXE advertisement for NetBiosName: NetBiosName: {netBiosName}, Resource ID: {list[0]}");
		if (resultObject["StatusCode"].IntegerValue is not 0)
		{
			throw new SmsException(resultObject["Description"].StringValue);
		}

		return true; // Even if the Count==0 we send true, it's just that there was NO PXE Requests for the device!
	}

	/// <summary>Update an existing computer.</summary>
	/// <param name="connection">The connection.</param>
	/// <param name="netBiosName">Net BIOS name.</param>
	/// <param name="macAddress">MAC Address.</param>
	/// <returns>The <see cref="int"/>.</returns>
	/// <exception cref="ArgumentNullException">The System Management BIOS Name of MAC Address must be defined.</exception>
	public static int UpdateExistingComputer(WqlConnectionManager connection, string netBiosName, string macAddress)
	{
		try
		{
			Dictionary<string, object> inParams = new() { { "NetbiosName", netBiosName }, { "MACAddress", macAddress }, { "OverwriteExistingRecord", true } };
			IResultObject outParams = connection.ExecuteMethod("SMS_Site", "ImportMachineEntry", inParams);
			return outParams["ResourceID"].IntegerValue;
		}
		catch (SmsException ex)
		{
			throw new SmsException($"{ex.Message}, {ex.Details}");
		}
	}
}