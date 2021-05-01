using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KmsTest
{
	public partial class Form1 : Form
	{
		private string SystemVerID = "";
		private string VersionName = "";

		public Form1()
		{
			InitializeComponent();
		}

		public enum SLIDTYPE
		{
			SL_ID_APPLICATION,
			SL_ID_PRODUCT_SKU,
			SL_ID_LICENSE_FILE,
			SL_ID_LICENSE,
			SL_ID_PKEY,
			SL_ID_ALL_LICENSES,
			SL_ID_ALL_LICENSE_FILES,
			SL_ID_STORE_TOKEN,
			SL_ID_LAST
		}
		public enum SLDATATYPE
		{
			SL_DATA_NONE,
			SL_DATA_SZ,
			SL_DATA_DWORD,
			SL_DATA_BINARY,
			SL_DATA_MULTI_SZ,
			SL_DATA_SUM
		}
		public struct ChannelEnum
		{
			public const string eRtail = "Retail";
			public const string eOEM = "OEM";
			public const string eVolume = "Volume";
			public const string eMAK = "Volume:MAK";
			public const string eGVLK = "Volume:GVLK";
		}

		[StructLayout(LayoutKind.Sequential)]
		public class RtlOsVersionInfoExW
		{
			public UInt32 dwOSVersionInfoSize;
			public UInt32 dwMajorVersion;
			public UInt32 dwMinorVersion;
			public UInt32 dwBuildNumber;
			public UInt32 dwPlataformId;
			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
			public string szCSDVersion;
			public UInt16 wServicePackMajor;
			public UInt16 wServicePackMinor;
			public UInt16 wSuiteMask;
			public byte bProductType;
			public byte bReserved;
		}

		[DllImport("pkeyhelper.dll", EntryPoint = "SkuGetProductKeyForEdition", CharSet = CharSet.Auto)]
		public static extern int SkuGetProductKeyForEdition(int EditionId, string Channel, ref IntPtr pProductKey, ref IntPtr pPfn);

		[DllImport("sppc.dll", EntryPoint = "SLOpen", ExactSpelling = false, CharSet = CharSet.Unicode)]
		public extern static int SLOpen(ref IntPtr hSLC);

		[DllImport("sppc.dll", EntryPoint = "SLClose", ExactSpelling = false, CharSet = CharSet.Unicode)]
		public extern static int SLClose(IntPtr hSLC);

		[DllImport("slc.dll", EntryPoint = "SLGetWindowsInformationDWORD", ExactSpelling = false, CharSet = CharSet.Unicode)]
		public static extern int SLGetWindowsInformationDWORD(string ValueName, ref int Value);

		[DllImport("sppc.dll", EntryPoint = "SLInstallProofOfPurchase", ExactSpelling = false, CharSet = CharSet.Unicode)]
		public extern static int SLInstallProofOfPurchase(IntPtr hSLC, string pwszPKeyAlgorithm, string pwszPKeyString, uint cbPKeySpecificData, byte[] pbPKeySpecificData, ref Guid PKeyId);

		[DllImport("sppc.dll", EntryPoint = "SLGetPKeyInformation", ExactSpelling = false, CharSet = CharSet.Unicode)]
		public extern static int SLGetPKeyInformation(IntPtr hOSLC, ref Guid pPKeyId, string pwszValueName, ref SLDATATYPE peDataType, ref uint pcbValue, ref IntPtr ppbValue);

		[DllImport("sppcext.dll", EntryPoint = "SLActivateProduct", ExactSpelling = false, CharSet = CharSet.Unicode)]
		public extern static int SLActivateProduct(IntPtr hSLC, ref Guid pProductSkuId, byte[] cbAppSpecificData, byte[] pvAppSpecificData, byte[] pActivationInfo, string pwszProxyServer, ushort wProxyPort);

		[DllImport("osppc.dll", EntryPoint = "SLOpen", ExactSpelling = false, CharSet = CharSet.Unicode)]
		public extern static int SLOpenO(ref IntPtr hSLC);

		[DllImport("osppc.dll", EntryPoint = "SLClose", ExactSpelling = false, CharSet = CharSet.Unicode)]
		public extern static int SLCloseO(IntPtr hSLC);

		[DllImport("osppc.dll", EntryPoint = "SLInstallProofOfPurchase", ExactSpelling = false, CharSet = CharSet.Unicode)]
		public extern static int SLInstallProofOfPurchaseO(IntPtr hSLC, string pwszPKeyAlgorithm, string pwszPKeyString, uint cbPKeySpecificData, byte[] pbPKeySpecificData, ref Guid PKeyId);

		[DllImport("osppc.dll", EntryPoint = "SLGetPKeyInformation", ExactSpelling = false, CharSet = CharSet.Unicode)]
		public extern static int SLGetPKeyInformationO(IntPtr hOSLC, ref Guid pPKeyId, string pwszValueName, ref SLDATATYPE peDataType, ref uint pcbValue, ref IntPtr ppbValue);

		[DllImport("osppcext.dll", EntryPoint = "SLActivateProduct", ExactSpelling = false, CharSet = CharSet.Unicode)]
		public extern static int SLActivateProductO(IntPtr hSLC, ref Guid pProductSkuId, byte[] cbAppSpecificData, byte[] pvAppSpecificData, byte[] pActivationInfo, string pwszProxyServer, ushort wProxyPort);


		[DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
		public extern static bool SetDllDirectory(string lpPathName);
		[DllImport("kernel32", SetLastError = true)]
		public extern static IntPtr LoadLibrary(string lpFileName);

		[DllImport("Ntdll.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
		public extern static int RtlGetVersion([In(), Out()] RtlOsVersionInfoExW osversion);
		public static Version NativeOsVersion()
		{
			RtlOsVersionInfoExW osVersionInfo = new RtlOsVersionInfoExW();
			osVersionInfo.dwOSVersionInfoSize = (uint)Marshal.SizeOf(typeof(RtlOsVersionInfoExW));
			int status = RtlGetVersion(osVersionInfo);
			if (status != 0)
			{
				return Environment.OSVersion.Version;
			}

			return new Version((int)osVersionInfo.dwMajorVersion, (int)osVersionInfo.dwMinorVersion, (int)osVersionInfo.dwBuildNumber);
		}
		private void ExtractEmbeddedResource(string outputDir, string resourceLocation, List<string> files)
		{
			foreach (string filePath in files)
			{
				if (File.Exists(Path.Combine(outputDir, filePath.Replace("32", "").Replace("64", "")))) continue;
				using (System.IO.Stream stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceLocation + "." + filePath))
				{
					using (FileStream fileStream = new FileStream(Path.Combine(outputDir, filePath.Replace("32", "").Replace("64", "")), System.IO.FileMode.Create))
					{
						for (int i = 0; i < stream.Length; i++)
						{
							fileStream.WriteByte((byte)stream.ReadByte());
						}
						fileStream.Close();
					}

				}
			}
		}
		private static string GetOsppcPath()
		{
			string OsppcPath = "";
			try
			{
				var registryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, (Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32)).OpenSubKey("SOFTWARE\\Microsoft\\OfficeSoftwareProtectionPlatform");
				if (registryKey != null)
				{
					OsppcPath = registryKey.GetValue("Path").ToString();
				}
			}
			catch { }
			return OsppcPath;
		}
		public static bool LoadOSPPC(string dllname)
		{
			var osppPath = GetOsppcPath();
			if (osppPath != "")
			{
				osppPath = osppPath + dllname;
				if (File.Exists(osppPath))
				{
					SetDllDirectory(osppPath);
					IntPtr dllptr = LoadLibrary(osppPath);
					if (dllptr == IntPtr.Zero)
					{
						Console.WriteLine(Marshal.GetLastWin32Error());
						osppPath = Environment.ExpandEnvironmentVariables("%ProgramFiles(x86)%") + "\\Common Files\\microsoft shared\\OfficeSoftwareProtectionPlatform\\" + dllname;
						if (File.Exists(osppPath))
						{
							dllptr = LoadLibrary(osppPath);
							if (dllptr == IntPtr.Zero)
							{
								return false;
							}
						}
						else
						{
							return false;
						}
					}
					else
					{
						return true;
					}
				}
			}
			return false;
		}
		private void AddExclusions2WindowsDefender(string ExcPath)
		{
			string sComputerID = null;
			try
			{
				ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\Microsoft\\Windows\\Defender", "SELECT * FROM MSFT_MpPreference");
				foreach (ManagementObject queryObj in searcher.Get())
				{
					sComputerID = queryObj["ComputerID"].ToString();
				}
			}
			catch (ManagementException ex)
			{
				Console.WriteLine("An error occurred while querying for WMI data: " + ex.Message);
			}

			try
			{
				string sPath = "MSFT_MpPreference.ComputerID='" + sComputerID + "'";
				ManagementObject classInstance = new ManagementObject("root\\Microsoft\\Windows\\Defender", sPath, null);
				ManagementBaseObject inParams = classInstance.GetMethodParameters("Add");
				string[] arrExclusionPath = new string[2];
				arrExclusionPath[0] = ExcPath;
				inParams["ExclusionPath"] = arrExclusionPath;
				ManagementBaseObject outParams = classInstance.InvokeMethod("Add", inParams, null);
			}
			catch (ManagementException ex)
			{
				Console.WriteLine("An error occurred while trying to execute the WMI method: " + ex.Message);
			}
		}
		private void Form1_Load(object sender, EventArgs e)
		{
			string OsType = (string)Registry.GetValue("HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Environment", "PROCESSOR_ARCHITECTURE", "Key does Not exist");
			if (OsType == "x86")
			{
				if (File.Exists(Environment.GetEnvironmentVariable("windir") + "\\System32\\SppExtComObjHook.dll") == false)
				{
					AddExclusions2WindowsDefender(Environment.GetEnvironmentVariable("windir") + "\\System32\\SppExtComObjHook.dll");
					ExtractEmbeddedResource(Environment.GetEnvironmentVariable("windir") + "\\System32", System.Reflection.Assembly.GetExecutingAssembly().GetName().Name.Replace("-", "_"), new List<string>() { "SppExtComObjHook32.dll" });
				}
			}
			else
			{
				if (File.Exists(Environment.GetEnvironmentVariable("windir") + "\\SysNative\\SppExtComObjHook.dll") == false)
				{
					AddExclusions2WindowsDefender(Environment.GetEnvironmentVariable("windir") + "\\SysNative\\SppExtComObjHook.dll");
					ExtractEmbeddedResource(Environment.GetEnvironmentVariable("windir") + "\\SysNative", System.Reflection.Assembly.GetExecutingAssembly().GetName().Name.Replace("-", "_"), new List<string>() { "SppExtComObjHook64.dll" });
				}
			}
		}
		private void btn_KMS_Click(object sender, EventArgs e)
		{
			string productKey = "";
			int EditionId = 0;
			int szResult = SLGetWindowsInformationDWORD("Kernel-ProductInfo", ref EditionId);
			if (szResult == 0)
			{
				IntPtr RetProductKey = new IntPtr();
				IntPtr RetPfn = new IntPtr();
				szResult = SkuGetProductKeyForEdition(EditionId, ChannelEnum.eGVLK, ref RetProductKey, ref RetPfn);
				if (szResult == 0)
				{
					productKey = Marshal.PtrToStringUni(RetProductKey);
				}
			}

			Guid pKeyId = new Guid();
			string nErrorcode = InstallProductKey(productKey, ref pKeyId);
			if (string.IsNullOrEmpty(nErrorcode))
			{
				AddKMSInfor();
				nErrorcode = ActivateProductKey(pKeyId);
				if (string.IsNullOrEmpty(nErrorcode) == false)
				{
					Console.WriteLine("KMS activation failed, error code: " + nErrorcode);
				}
				else
				{
					Console.WriteLine("KMS activation succeeded! Please refresh the license.");
				}
			}
			else
			{
				Console.WriteLine("Installtion productKey failed, error code: " + nErrorcode);
			}

		}
		private string InstallProductKey(string ProductKey, ref Guid pKeyId)
		{
			SystemVerID = NativeOsVersion().Major.ToString() + "." + NativeOsVersion().Minor.ToString();
			string nErrorcode = string.Empty;
			IntPtr hSLC = new IntPtr();
			var hResult = SLOpen(ref hSLC);
			if (hResult == 0)
			{
				if (Convert.ToDouble(SystemVerID) < 6.2)
				{
					hResult = SLInstallProofOfPurchase(hSLC, "msft:rm/algorithm/pkey/2005", ProductKey, 0, null, ref pKeyId);
				}
				else
				{
					hResult = SLInstallProofOfPurchase(hSLC, "msft:rm/algorithm/pkey/2009", ProductKey, 0, null, ref pKeyId);
				}
				if (hResult != 0)
				{
					nErrorcode = "0x" + hResult.ToString("X8");
					Console.WriteLine("Installation key failed with error code:" + nErrorcode);
				}
			}
			else
			{
				nErrorcode = "0x" + hResult.ToString("X8");
				Console.WriteLine("Installation key failed with error code:" + nErrorcode);
			}
			SLClose(hSLC);
			return nErrorcode;
		}
		private string ActivateProductKey(Guid pKeyId)
		{
			string nErrorcode = string.Empty;
			IntPtr hSLC = new IntPtr();
			var hResult = SLOpen(ref hSLC);
			uint pcbValue = 0;
			IntPtr ppbValue = System.IntPtr.Zero;
			SLDATATYPE peDataType = new SLDATATYPE();
			hResult = SLGetPKeyInformation(hSLC, ref pKeyId, "ProductSkuId", ref peDataType, ref pcbValue, ref ppbValue);
			if (hResult == 0)
			{
				byte[] bSkuId = new byte[pcbValue];
				Marshal.Copy(ppbValue, bSkuId, 0, (int)pcbValue);
				Guid pSkuId = new Guid(bSkuId);
				hResult = SLActivateProduct(hSLC, ref pSkuId, null, null, null, null, 0);
				if (hResult == 0)
				{
					SLClose(hSLC);
					Console.WriteLine("KMS activation succeeded! Please refresh the license.");
				}
				else if (hResult == unchecked((int)0xC004F074))
				{
					nErrorcode = "0x" + hResult.ToString("X8");
					Console.WriteLine("KMS activation failed, error code: " + nErrorcode);
				}
			}
			else
			{
				nErrorcode = "0x" + hResult.ToString("X8");
				Console.WriteLine("KMS activation failed, error code: " + nErrorcode);
			}
			SLClose(hSLC);
			return nErrorcode;
		}
		private bool InstallProductKeyO(string ProductKey)
		{
			var nErrorcode = "";
			try
			{
				using (ManagementObjectSearcher query = new ManagementObjectSearcher("SELECT Version FROM OfficeSoftwareProtectionService"))
				{
					foreach (ManagementObject product in query.Get())
					{
						uint result = 1;
						try
						{
							result = Convert.ToUInt32(product.InvokeMethod("InstallProductKey", new[] { ProductKey }));
						}
						catch (COMException ex)
						{
							int hr = System.Runtime.InteropServices.Marshal.GetHRForException(ex);
							nErrorcode = "0x" + hr.ToString("X8");
							Console.WriteLine("Installation failed, Error Code:" + nErrorcode);
							return false;
						}
					}
				}
			}
			catch (Exception ex)
			{
				int hr = System.Runtime.InteropServices.Marshal.GetHRForException(ex);
				nErrorcode = "0x" + hr.ToString("X8");
				Console.WriteLine("Installation failed, Error Code:" + nErrorcode);
				return false;
			}
			return true;

		}
		private string ActivateProductKeyO(string ProductKey)
		{
			string nErrorcode = "";
			ManagementObjectSearcher searcherObj = new ManagementObjectSearcher("root\\CIMV2", "SELECT Name,ID,PartialProductKey,OfflineInstallationId FROM OfficeSoftwareProtectionProduct WHERE PartialProductKey like '" + ProductKey.Substring(ProductKey.Trim().Length - 5, 5) + "' ");
			foreach (ManagementObject product in searcherObj.Get())
			{
				try
				{
					var results = Convert.ToUInt32(product.InvokeMethod("Activate", null));
					Console.WriteLine("KMS activation succeeded! Please refresh the license.");
				}
				catch (COMException ex)
				{
					Console.WriteLine("0x{0:x}", Marshal.GetHRForException(ex));
					int hr = System.Runtime.InteropServices.Marshal.GetHRForException(ex);
					nErrorcode = "0x" + hr.ToString("X8");
					Console.WriteLine("KMS activation failed, error code:" + nErrorcode);
				}
			}
			return nErrorcode;
		}
		public bool InstallOfficeProductKey(string ProductKeys)
		{
			if (LoadOSPPC("osppc.dll") == false)
			{
				return InstallProductKeyO(ProductKeys);
			}
			string nErrorcode = string.Empty;
			Guid pKeyId = new Guid();
			string PKeyAlgorithm = "msft:rm/algorithm/pkey/2009";
			try
			{
				IntPtr hOSLC = new IntPtr();
				var hResult = SLOpenO(ref hOSLC);
				if (hResult == 0)
				{
					hResult = SLInstallProofOfPurchaseO(hOSLC, PKeyAlgorithm, ProductKeys, 0, null, ref pKeyId);
					if (hResult == unchecked((int)0xC004F016))
					{
						PKeyAlgorithm = "msft:rm/algorithm/pkey/2005";
						hResult = SLInstallProofOfPurchaseO(hOSLC, PKeyAlgorithm, ProductKeys, 0, null, ref pKeyId);
					}
					if (hResult == 0)
					{
						//ActivateOfficeProductKey(hOSLC, pKeyId, ProductKeys, PKeyAlgorithm);
					}
					else
					{
						nErrorcode = "0x" + hResult.ToString("X8");
						Console.WriteLine("Installation key failed with error code:" + nErrorcode);
					}
				}
			}
			catch (Exception ex)
			{
				int hr = System.Runtime.InteropServices.Marshal.GetHRForException(ex);
				nErrorcode = "0x" + hr.ToString("X8");
				return false;
			}

			return true;
		}
		public string ActivateOfficeProductKey(IntPtr hOSLC, string productKey, Guid pKeyId)
		{
			string nErrorcode = string.Empty;
			if (LoadOSPPC("OSPPCEXT.dll") == false)
			{
				return ActivateProductKeyO(productKey);
			}

			uint pcbValue = 0;
			IntPtr ppbValue = System.IntPtr.Zero;
			SLDATATYPE peDataType = new SLDATATYPE();
			int hResult = SLGetPKeyInformationO(hOSLC, ref pKeyId, "ProductSkuId", ref peDataType, ref pcbValue, ref ppbValue);
			if (hResult == 0)
			{
				byte[] bSkuId = new byte[pcbValue];
				Marshal.Copy(ppbValue, bSkuId, 0, (int)pcbValue);
				Guid pSkuId = new Guid(bSkuId);
				hResult = SLActivateProductO(hOSLC, ref pSkuId, null, null, null, null, 0);
				if (hResult == 0)
				{
					SLClose(hOSLC);
					Console.WriteLine("KMS activation succeeded! Please refresh the license.");
				}
				else if (hResult == unchecked((int)0xC004F074))
				{
					nErrorcode = "0x" + hResult.ToString("X8");
					Console.WriteLine("KMS activation failed, error code: " + nErrorcode);
				}
			}
			else
			{
				nErrorcode = "0x" + hResult.ToString("X8");
				Console.WriteLine("KMS activation failed, error code: " + nErrorcode);
			}
			SLClose(hOSLC);
			return nErrorcode;
		}
		private bool AddKMSInfor()
		{
			string nErrorcode = "";
			ManagementObjectSearcher query = new ManagementObjectSearcher();
			if (Convert.ToDouble(SystemVerID) < 6.2 && VersionName.Contains("Win") == false || VersionName.Contains("2010"))
			{
				query = new ManagementObjectSearcher("SELECT Version FROM OfficeSoftwareProtectionService");
			}
			else
			{
				query = new ManagementObjectSearcher("SELECT Version FROM SoftwareLicensingService");
			}
			foreach (ManagementObject item in query.Get())
			{
				try
				{
					if (Convert.ToDouble(SystemVerID) < 6.2 && VersionName.Contains("Win") == false || VersionName.Contains("2010"))
					{
						var result = Convert.ToUInt32(item.InvokeMethod("SetKeyManagementServiceMachine", new[] { "0.0.0.0" }));
					}
					else
					{
						var result = Convert.ToUInt32(item.InvokeMethod("SetKeyManagementServiceMachine", new[] { "0.0.0.0" })); //172.0.0.2
					}
					item.InvokeMethod("SetKeyManagementServicePort", new[] { "1688" });
				}
				catch (COMException ex)
				{
					int hr = System.Runtime.InteropServices.Marshal.GetHRForException(ex);
					nErrorcode = "0x" + hr.ToString("X8");
					Console.WriteLine("Adding KMS server port failed, error code:" + nErrorcode);
					return false;
				}
			}
			return true;
		}
		private void SetRegistry()
		{
			RegistryKey OpenLMKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, (Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32));
			RegistrySecurity rs = new RegistrySecurity();
			string currentUserStr = Environment.UserDomainName + "\\" + Environment.UserName;
			rs.AddAccessRule(new RegistryAccessRule(currentUserStr, RegistryRights.WriteKey | RegistryRights.ReadKey | RegistryRights.Delete | RegistryRights.FullControl, AccessControlType.Allow));

			RegistryKey regkey = null;
			if (Convert.ToDouble(SystemVerID) < 6.2 && VersionName.Contains("Win") == false || VersionName.Contains("2010"))
			{
				regkey = OpenLMKey.OpenSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\osppsvc.exe", true);
				if (regkey == null)
				{
					using (var subRegKey = OpenLMKey.CreateSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\osppsvc.exe", RegistryKeyPermissionCheck.Default))
					{
						subRegKey.SetAccessControl(rs);
						subRegKey.SetValue("KMS_ActivationInterval", 43200, RegistryValueKind.DWord);
						subRegKey.SetValue("KMS_Emulation", 1, RegistryValueKind.DWord);
						subRegKey.SetValue("KMS_RenewalInterval", 43200, RegistryValueKind.DWord);
						subRegKey.SetValue("VerifierDlls", "SppExtComObjHook.dll", RegistryValueKind.String);
						subRegKey.SetValue("VerifierDebug", 0, RegistryValueKind.DWord);
						subRegKey.SetValue("VerifierFlags", unchecked((int)0x80000000), RegistryValueKind.DWord);
						subRegKey.SetValue("GlobalFlag", 0x100, RegistryValueKind.DWord);
					}
				}
				else
				{
					regkey.SetAccessControl(rs);
					regkey.SetValue("KMS_ActivationInterval", 43200, RegistryValueKind.DWord);
					regkey.SetValue("KMS_Emulation", 1, RegistryValueKind.DWord);
					regkey.SetValue("KMS_RenewalInterval", 43200, RegistryValueKind.DWord);
					regkey.SetValue("VerifierDlls", "SppExtComObjHook.dll", RegistryValueKind.String);
					regkey.SetValue("VerifierDebug", 0, RegistryValueKind.DWord);
					regkey.SetValue("VerifierFlags", unchecked((int)0x80000000), RegistryValueKind.DWord);
					regkey.SetValue("GlobalFlag", 0x100, RegistryValueKind.DWord);
				}

				regkey = OpenLMKey.OpenSubKey("SOFTWARE\\Microsoft\\OfficeSoftwareProtectionPlatform", true);
				if (regkey == null)
				{
					using (var subRegKey = OpenLMKey.CreateSubKey("SOFTWARE\\Microsoft\\OfficeSoftwareProtectionPlatform", RegistryKeyPermissionCheck.Default))
					{
						regkey.SetAccessControl(rs);
						regkey.SetValue("KeyManagementServiceName", "0.0.0.0", RegistryValueKind.String);
					}
				}
				else
				{
					regkey.SetAccessControl(rs);
					regkey.SetValue("KeyManagementServiceName", "0.0.0.0", RegistryValueKind.String);
				}
			}
			else
			{
				regkey = OpenLMKey.OpenSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\sppsvc.exe", true);
				if (regkey == null)
				{
					using (var subRegKey = OpenLMKey.CreateSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\sppsvc.exe", RegistryKeyPermissionCheck.Default))
					{
						subRegKey.SetAccessControl(rs);
						subRegKey.SetValue("KMS_ActivationInterval", 43200, RegistryValueKind.DWord);
						subRegKey.SetValue("KMS_Emulation", 1, RegistryValueKind.DWord);
						subRegKey.SetValue("KMS_RenewalInterval", 43200, RegistryValueKind.DWord);
						subRegKey.SetValue("VerifierDlls", "SppExtComObjHook.dll", RegistryValueKind.String);
						subRegKey.SetValue("VerifierDebug", 0, RegistryValueKind.DWord);
						subRegKey.SetValue("VerifierFlags", unchecked((int)0x80000000), RegistryValueKind.DWord);
						subRegKey.SetValue("GlobalFlag", 0x100, RegistryValueKind.DWord);
					}
				}
				else
				{
					regkey.SetAccessControl(rs);
					regkey.SetValue("KMS_ActivationInterval", 43200, RegistryValueKind.DWord);
					regkey.SetValue("KMS_Emulation", 1, RegistryValueKind.DWord);
					regkey.SetValue("KMS_RenewalInterval", 43200, RegistryValueKind.DWord);
					regkey.SetValue("VerifierDlls", "SppExtComObjHook.dll", RegistryValueKind.String);
					regkey.SetValue("VerifierDebug", 0, RegistryValueKind.DWord);
					regkey.SetValue("VerifierFlags", unchecked((int)0x80000000), RegistryValueKind.DWord);
					regkey.SetValue("GlobalFlag", 0x100, RegistryValueKind.DWord);
				}

				if (SystemVerID.Contains("6.2"))
				{
					regkey = OpenLMKey.OpenSubKey("SOFTWARE\\Policies\\Microsoft\\Windows NT\\CurrentVersion\\Software Protection Platform\\", true);
					if (regkey == null)
					{
						using (var subRegKey = OpenLMKey.CreateSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\SoftwareProtectionPlatform", RegistryKeyPermissionCheck.Default))
						{
							regkey.SetValue("NoGenTicket", 1, RegistryValueKind.DWord); //NoAcquireGT
						}
					}
					else
					{
						regkey.SetValue("NoGenTicket", 1, RegistryValueKind.DWord);
					}
				}
			}

			regkey = OpenLMKey.OpenSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\SppExtComObj.exe", true);
			if (regkey == null)
			{
				using (var subRegKey = OpenLMKey.CreateSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Image File Execution Options\\SppExtComObj.exe", RegistryKeyPermissionCheck.Default))
				{
					subRegKey.SetAccessControl(rs);
					subRegKey.SetValue("VerifierDlls", "SppExtComObjHook.dll", RegistryValueKind.String);
					subRegKey.SetValue("VerifierDebug", 0, RegistryValueKind.DWord);
					subRegKey.SetValue("VerifierFlags", unchecked((int)0x80000000), RegistryValueKind.DWord);
					subRegKey.SetValue("GlobalFlag", 0x100, RegistryValueKind.DWord);
					subRegKey.SetValue("KMS_ActivationInterval", 43200, RegistryValueKind.DWord);
					subRegKey.SetValue("KMS_Emulation", 1, RegistryValueKind.DWord);
					subRegKey.SetValue("KMS_HWID", 4187226795851251830, RegistryValueKind.QWord);
					subRegKey.SetValue("KMS_RenewalInterval", 43200, RegistryValueKind.DWord);
				}
			}
			else
			{
				regkey.SetAccessControl(rs);
				regkey.SetValue("VerifierDlls", "SppExtComObjHook.dll", RegistryValueKind.String);
				regkey.SetValue("VerifierDebug", 0, RegistryValueKind.DWord);
				regkey.SetValue("VerifierFlags", unchecked((int)0x80000000), RegistryValueKind.DWord);
				regkey.SetValue("GlobalFlag", 0x100, RegistryValueKind.DWord);
				regkey.SetValue("KMS_ActivationInterval", 43200, RegistryValueKind.DWord);
				regkey.SetValue("KMS_Emulation", 1, RegistryValueKind.DWord);
				regkey.SetValue("KMS_HWID", 4187226795851251830, RegistryValueKind.QWord);
				regkey.SetValue("KMS_RenewalInterval", 43200, RegistryValueKind.DWord);
			}

		}


	}
}
