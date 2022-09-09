namespace SccmComInterop;

public static partial class CmInterop
{
	/// <summary>The folder type.</summary>
	public enum FolderType
	{
		/// <summary>The SMS package.</summary>
		SMS_Package = 2,

		/// <summary>The SMS advertisement.</summary>
		SMS_Advertisement = 3,

		/// <summary>The SMS query.</summary>
		SMS_Query = 7,

		/// <summary>The SMS report.</summary>
		SMS_Report = 8,

		/// <summary>The SMS metered product rule.</summary>
		SMS_MeteredProductRule = 9,

		/// <summary>The SMS configuration item.</summary>
		SMS_ConfigurationItem = 11,

		/// <summary>The SMS operating system install package.</summary>
		SMS_OperatingSystemInstallPackage = 14,

		/// <summary>The SMS state migration.</summary>
		SMS_StateMigration = 17,

		/// <summary>The SMS image package.</summary>
		SMS_ImagePackage = 18,

		/// <summary>The SMS boot image package.</summary>
		SMS_BootImagePackage = 19,

		/// <summary>The SMS task sequence package.</summary>
		SMS_TaskSequencePackage = 20,

		/// <summary>The SMS device setting package.</summary>
		SMS_DeviceSettingPackage = 21,

		/// <summary>The SMS driver package.</summary>
		SMS_DriverPackage = 23,

		/// <summary>The SMS driver.</summary>
		SMS_Driver = 25,

		/// <summary>The SMS software update.</summary>
		SMS_SoftwareUpdate = 1011,

		/// <summary>SMS Collection of Devices</summary>
		SMS_Collection_Device = 5000
	}
}