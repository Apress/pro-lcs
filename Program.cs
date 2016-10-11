using System;
using System.Collections.Generic;
using System.Text;
using CommandLine.Utility;

namespace sipchange
{
    using System.Management;
    using System.Collections;
    using System.Collections.Specialized;
    using System.DirectoryServices;
    using System.Text.RegularExpressions;

    class Program
    {
        static string sQuery;
        static string sQueryContacts;
        static string SourceDomainURI;
        static string TargetDomainURI;

        public Program()
        {
            // query string to return instance ID of user enabled for LCS using WMI
            sQuery = "select * from MSFT_SIPESUserSetting where UserDN = '";
            // query string to return all contacts given a user's instance ID
            sQueryContacts = "select * from MSFT_SIPESUserContactData where UserInstanceID = '";
        }

        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("\ncommand-line arguments are required.");
                return;
            }

            Arguments CommandLine = new Arguments(args);

            if (CommandLine["?"] != null)
            {
                Console.WriteLine("\ncommand-line arguments:");
                Console.WriteLine("\t/source:<name>\t\t- SIP domain to match");
                Console.WriteLine("\t/target:<name>\t\t- SIP domain to change");
                return;
            }
            if (CommandLine["source"] != null)
            {
                SourceDomainURI = CommandLine["source"];
            }
            if (CommandLine["target"] != null)
            {
                TargetDomainURI = CommandLine["target"];
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
				return;
			}

            // Instantiate class.
            Program SipDomain = new Program();

            string userDN = null;
            try
            {
                DirectorySearcher Users = new DirectorySearcher(deRoot, "(&(objectCategory=person)(objectCategory=user)(msRTCSIP-UserEnabled=TRUE))");

                foreach(SearchResult user in Users.FindAll())
                {
                    // Obtain DN of user by removing the first 7 characters "LDAP://".
                    userDN = user.Path.Substring(7);
                    SipDomain.ChangeSIPdomain(userDN);
                    Console.WriteLine();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return;
            }
        }

        public void ChangeSIPdomain(string userDN)
        {
            try
            {
                // Query for user's instance ID
                ManagementObjectSearcher oSearcher = new ManagementObjectSearcher(sQuery + userDN + "'");
                ManagementObjectCollection oCollection = oSearcher.Get();

                foreach (ManagementObject user in oCollection)
                {
                    Console.Write(user["PrimaryURI"].ToString());

                    if(Regex.IsMatch(user["PrimaryURI"].ToString(), SourceDomainURI))
                    {
                        // Modify user SIP URI domain portion
                        string URI = Regex.Replace(user["PrimaryURI"].ToString(), @"@[-\w.]+", "@" + TargetDomainURI);
                        Console.Write(" -> " + URI);
                        user["PrimaryURI"] = URI;
                        // Commit changes
                        //user.Put();
                    }
                    Console.WriteLine();

                    // Check whether user is unassigned
                    if (user["HomeServerDN"] == null)
                    {
                        Console.WriteLine("WARNING: this user is unassigned");
                        // If the user is unassigned, then it won't be possible to read their contact list
                        return;
                    }

                    UpdateContactsURI(user["InstanceID"].ToString());
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("ChangeSIPdomain(): " + e.Message);
            }
        }

        private void UpdateContactsURI(string InstanceID)
        {
            try
            {
                ManagementObjectSearcher oSearcher = new ManagementObjectSearcher(sQueryContacts + InstanceID + "'");
                ManagementObjectCollection oCollection = oSearcher.Get();

                // Check whether user has any contacts in their contact list
                if (oCollection.Count == 0)
                {
                    return;
                }

                foreach (ManagementObject contact in oCollection)
                {
                    if (Regex.IsMatch(contact.GetPropertyValue("SIPURI").ToString(), SourceDomainURI))
                    {
                        string URI = Regex.Replace(contact.GetPropertyValue("SIPURI").ToString(), @"@[-\w.]+", "@" + TargetDomainURI);
                        // Display only contacts that are modified
                        Console.WriteLine("\tcontact: " + contact.GetPropertyValue("SIPURI") + " -> " + URI);
                        contact.SetPropertyValue("SIPURI", URI);
                        // Commit change
                        //contact.Put();
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("WARNING: unable to contact user\'s home server to read contact list");
            }
        }
    }
}
