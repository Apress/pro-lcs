#define LIMIT

namespace lcsutil 
{
	using System;
	using System.Management;
	using System.Collections;
	using System.Collections.Specialized;
	using System.DirectoryServices;
	using System.Text.RegularExpressions;
	using CommandLine.Utility;

	/// <summary>
	/// Summary description for BulkConfigUsers.
	/// </summary>
	class BulkConfigUsers
	{
		enum cmd {LCS, FED, REMOTE, PIC, ARCHIVE, RCC, MEETING};
        static string sQuery;
		static bool LcsValue, FederationValue, RemoteAccessValue, PicValue;
		private
			string AppendURI = null;
			string ServerDN = null;
            string domainURI = null;
			static string ArchiveValue = null;

        public BulkConfigUsers()
        {
            // string query to return instance ID of user enabled for LCS using WMI
            sQuery = "select * from MSFT_SIPESUserSetting where UserDN = '";
        }

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static int Main(string[] args)
		{
			if(args.Length == 0)
			{
				Console.WriteLine("\nNo arguments were provided.");
				(new BulkConfigUsers()).Usage();
				return 0;
			}

			// Bit field to store list of command operations to perform.
			BitArray Op = new BitArray(7, false);
			// store list of users to configure.
			Array users = null;
			// Active Directory group to configure.
			string group = null;
			// Active Directory DN to search users to configure.
			string containerDN = null;

			Arguments CommandLine = new Arguments(args);

            // Instantiate class.
			BulkConfigUsers BulkConfig = new BulkConfigUsers();
			
            if(CommandLine["?"] != null)
            {
                BulkConfig.Usage();
                return 0;
            }
            if(CommandLine["users"] != null)
            {
                char[] separator = {','};
                users = CommandLine["users"].Split(separator);
            }
            if(CommandLine["group"] != null)
            {
                group = CommandLine["group"];
				Console.WriteLine("group name: '" + group+ "'");
            }
			if(CommandLine["container"] != null)
			{
				containerDN = CommandLine["container"];
			}
			if(CommandLine["append"] != null)
			{
				BulkConfig.SetUriAppendString(CommandLine["append"]);
			}
			if(CommandLine["server"] != null)
            {
                bool bResult = BulkConfig.SetHomeServer(CommandLine["server"]);
            }
            if(CommandLine["domain"] != null)
            {
                BulkConfig.SetDomainURI(CommandLine["domain"]);
            }
            if(CommandLine["lcs"] != null)
            {
                Op.Set((int)cmd.LCS, true);
				if(CommandLine["lcs"] == "y")
					LcsValue = true;
				else if(CommandLine["lcs"] == "n")
					LcsValue = false;
				else
				{
					Console.WriteLine("\nMissing parameter: y | n");
					BulkConfig.Usage();
					return 0;
				}
            }
            if(CommandLine["fed"] != null)
            {
                Op.Set((int)cmd.FED, true);
				if(CommandLine["fed"] == "y")
					FederationValue = true;
				else if(CommandLine["fed"] == "n")
					FederationValue = false;
				else
				{
					Console.WriteLine("\nMissing parameter: y | n");
					BulkConfig.Usage();
					return 0;
				}
            }
            if(CommandLine["remote"] != null)
            {
                Op.Set((int)cmd.REMOTE, true);
                if(CommandLine["remote"] == "y")
                    RemoteAccessValue = true;
                else if(CommandLine["remote"] == "n")
                    RemoteAccessValue = false;
                else
				{
					Console.WriteLine("\nMissing parameter: y | n");
					BulkConfig.Usage();
					return 0;
				}
			}
			if(CommandLine["pic"] != null)
			{
				Op.Set((int)cmd.PIC, true);
				if(CommandLine["pic"] == "y")
					PicValue = true;
				else if(CommandLine["pic"] == "n")
					PicValue = false;
				else
				{
					Console.WriteLine("\nMissing parameter: y | n");
					BulkConfig.Usage();
					return 0;
				}
			}
			
			/// Validate command-line arguments.
			if(users == null && group == null && containerDN == null)
			{
				Console.WriteLine("\nMissing argument: list of users, group or container name.");
                BulkConfig.Usage();
				return 0;
			}

			Console.WriteLine();
#if LIMIT
			int count = 0; // used only for trial version
			int limit = 5; // trial version limit
#endif

			if(containerDN != null)
			{
				try
				{
					// Query all users under a container from a local GC since all SIP related 
					// information of users are marked for GC replication.
					DirectoryEntry deContainer = new DirectoryEntry("GC://" + containerDN);
					DirectorySearcher srchUsers = new DirectorySearcher(deContainer);
					srchUsers.SearchScope = SearchScope.Subtree;
					srchUsers.Filter = ("(&(objectCategory=person)(objectClass=user))");

					// Find all users under an Active Directory container.
					foreach(SearchResult user in srchUsers.FindAll())
                    {
						BulkConfig.ConfigureUser(user.Path.Substring(5), Op);
#if LIMIT
						count++;
						if(count > limit)
						{
							Console.WriteLine("You have reached the maximum usage of this trial version");
							break;
						}
#endif
					}
				}
				catch(Exception e)
				{
					Console.WriteLine(e.Message);
					return 0;
				}
				return 1;
			}

			DirectoryEntry deRoot = null;

			try
			{
				DirectoryEntry  deRootDSE = new DirectoryEntry("LDAP://RootDSE");
				string sRootDomain = "LDAP://" + deRootDSE.Properties["rootDomainNamingContext"].Value.ToString();
				deRoot = new DirectoryEntry(sRootDomain);
			}
			catch(Exception e)
			{
				Console.WriteLine("Failed to query Active Directory");
				Console.WriteLine(e.Message);
				return 0;
			}

			if(group != null)
			{
				string sGroup = null;

				try
				{
					// Query all users members of a group from a local GC since all SIP related 
					// information of users are marked for GC replication.
					DirectorySearcher srchGroup = new DirectorySearcher(deRoot);
					srchGroup.SearchScope = SearchScope.Subtree;
					srchGroup.Filter = ("(&(objectCategory=group)(name=" + group + "))");

					// Obtain DN of group.
					sGroup = srchGroup.FindOne().Path;
				}
				catch(Exception e)
				{
					Console.WriteLine(e.Message);
					return 0;
				}

				// Find all users member of group.
				DirectoryEntry deGroup = new DirectoryEntry(sGroup);
				foreach(object userDN in deGroup.Properties["member"] )
				{
					BulkConfig.ConfigureUser(userDN.ToString(), Op);
					count++;
					if(count > limit)
					{
						Console.WriteLine("You have reached the maximum usage of this trial version");
						break;
					}
				}
			}
			else if(users != null)
			{
				foreach(string user in users)
				{
					string userDN = null;
					try 
					{
						DirectorySearcher srchUser = new DirectorySearcher(deRoot,"(&(objectCategory=user)(cn=" + user + "))");

						// Obtain DN of user by removing the first 7 characters "LDAP://".
						userDN = srchUser.FindOne().Path.Substring(7);
					}
					catch(Exception e)
					{
						Console.WriteLine(e.Message);
						return 0;
					}

					BulkConfig.ConfigureUser(userDN, Op);
					count++;
					if(count > limit)
					{
						Console.WriteLine("You have reached the maximum usage of this trial version");
						break;
					}
				}
			}	

			return 1;
		}

        public bool SetHomeServer(string fqdn)
        {
            try
            {
                // Determine DN of homeserver
                string sQuery = "select * from msft_sippoolsetting where PoolFQDN='";
                ManagementObjectSearcher oSearcher = new ManagementObjectSearcher(sQuery + fqdn + "'"); 
            
                foreach(ManagementObject mo in oSearcher.Get())
                {
                    ServerDN = mo["PoolDN"].ToString();
                }
                return true;
            }
            catch(Exception e)
            {
                Console.WriteLine("\nUnable to query server DN: " + e.Message);
                return false;
            }
        }

		public bool SetUriAppendString(string append)
		{
			AppendURI = append;
			return true;
		}

        public bool SetDomainURI(string domain)
        {
            domainURI = String.Copy(domain);
			return true;
        }

		public void CreateLcsAccount(string userDN, BitArray Op)
		{
			Console.WriteLine("create user");

			// Use ADSI to access the user object
			DirectoryEntry deUser = new DirectoryEntry("LDAP://" + userDN);
			string URI = null;

            // Default is to use the email address as the user's SIP URI
            // if present; otherwise, use the SAM account name
			if(deUser.Properties["mail"].Value != null)
			{
				URI = deUser.Properties["mail"].Value.ToString();
			}
			else
			{
				URI = deUser.Properties["samAccountName"].Value.ToString();
			}
							
			ManagementClass mgmtClass = new ManagementClass("root/cimv2",
				"MSFT_SIPESUserSetting", new ObjectGetOptions());

			ManagementObject oUser = mgmtClass.CreateInstance();
			oUser["UserDN"] = userDN;
			oUser["DisplayName"] = null;	// this attribute must be null on creation
			oUser["PrimaryURI"] = "sip:" + URI;
			oUser["HomeServerDN"] = ServerDN;
			oUser["Enabled"] = true;

			this.Configure(oUser, Op);
		}

		public void Configure(ManagementObject user, BitArray Op)
		{
			if(Op.Get((int)cmd.LCS))
			{
				if(LcsValue == true && ServerDN == null && user["HomeServerDN"] == null)
				{
					// do nothing - can not enable a user for Live Communications Server 
					// if they are not homed on a pool
					Console.WriteLine("User can not enabled for LCS without specifying a server fqdn");
				}
				else
				{
					if(ServerDN != null)
					{
						user["HomeServerDN"] = ServerDN;
					}
					user["Enabled"] = LcsValue;
				}
			}
			else if(Op.Get((int)cmd.REMOTE))
			{
				user["EnabledForInternetAccess"] = RemoteAccessValue;
			}
			else if(Op.Get((int)cmd.FED))
			{
				user["EnabledForFederation"] = FederationValue;
			}
			else if(Op.Get((int)cmd.PIC))
			{
				user["PublicNetworkEnabled"] = PicValue;
			}
			else
			{
                if(domainURI != null)
                {
                    // Modify user SIP URI domain portion if one is provided
                    string URI = Regex.Replace(user["PrimaryURI"].ToString(), @"@[-\w.]+", "@"+domainURI);
                    user["PrimaryURI"] = URI;
                }
				else if(AppendURI != null)
				{
					// Modify the user SIP URI name portion by appending
					string URI = Regex.Replace(user["PrimaryURI"].ToString(), @"@", AppendURI+"@");
					user["PrimaryURI"] = URI;
				}
                else
                {
                    Console.WriteLine("Invalid configuration operation requested");
                }
			}
			try 
			{
				// Commit changes
				user.Put();
				Console.WriteLine(user["DisplayName"].ToString() + ": success");
			}
			catch(Exception e)
			{
				Console.WriteLine(e.Message);
				Console.WriteLine();
			}
		}

		public void ConfigureUser(string userDN, BitArray Op)
		{
			try
			{
				// Query for user's instance ID
				ManagementObjectSearcher oSearcher = new ManagementObjectSearcher(sQuery + userDN + "'"); 
				ManagementObjectCollection oCollection = oSearcher.Get();

				Console.WriteLine(userDN);
				int iCount = 0;

				foreach(ManagementObject mo in oCollection)
				{
					Configure(mo, Op);
					iCount++;
				}

				if(iCount == 0)
				{
					// User is not enabled for LCS; therefore, the LCS WMI provider
					// does not find the user in the LCS database.
					if(ServerDN != null)
					{
						CreateLcsAccount(userDN, Op);
					}
					else
					{
						Console.WriteLine(userDN + " can not be enabled for LCS");
					}
				}
			}
			catch(Exception e)
			{
				Console.WriteLine(e.Message);
				Console.WriteLine();
			}
		}

		public void Usage()
		{
			Console.WriteLine("\nLCS Utility");
			Console.WriteLine("\ncommand-line arguments:");
			Console.WriteLine("\t/users:<name>\t\t- list of users separated by commas");
			Console.WriteLine("\t/group:<name>\t\t- group name");
			Console.WriteLine("\t/container:<name>\t- DN");

			Console.WriteLine("\n\tModifiers:");
			Console.WriteLine("\t/append:<string>\t- append to user name portion of user's SIP URI");
			Console.WriteLine("\t/server:<fqdn>\t\t- server name (fqdn) to home users");
			Console.WriteLine("\t/domain:<string>\t- modify domain portion of user's SIP URI");

			Console.WriteLine("\n\tOperations:");
			Console.WriteLine("\t/lcs:y|n\t\t- enable or disable user for LCS");
			Console.WriteLine("\t/fed:y|n\t\t- enable or disable user for federation");
			Console.WriteLine("\t/remote:y|n\t\t- enable or disable user for remote access");
			Console.WriteLine("\t/pic:y|n\t\t- enable or disable user for public IM");
			Console.WriteLine();
		}
	}
}
