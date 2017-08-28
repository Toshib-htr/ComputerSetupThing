using System.Collections.Generic;
using System.Text;
using System.IO;
using System;
using System.Net.NetworkInformation;
using System.Management;
using System.Diagnostics;
using System.DirectoryServices.ActiveDirectory;
using Microsoft.Win32;
using WUApiLib;

namespace thing
{
	class Program
	{
		private static string CONFIG_FILE = "config.conf";
		private static string OUTPUT_FILE = "out.csv";

		static void Main(string[] args)
		{
			Dictionary<String, String> options = read_config();
			String name = "";

			string[] creds = new string[3];
			
			#region Active Directory Stuff
			if (options.ContainsKey("join_ADD") && options["join_ADD"] == "true")
			{
				Console.WriteLine("Joining AD Domain.... ");

				creds = join_AD_Domain((options.ContainsKey("AD_name") ? options["AD_name"] : ""),
						options.ContainsKey("AD_uname") ? options ["AD_uname"] : ""));

				Console.WriteLine("done\n");
			}
			#endregion

			#region Change the machine name
			if (options.ContainsKey("change_name") && options["change_name"] == "true" &&
			  options.ContainsKey("sys_name"))
			{
			    name = get_name(options["sys_name"], (options.ContainsKey("start_num") ? options["start_num"] : ""));
			    Console.Write("Changing machine name..... ");
			        changeMachineName(name, (creds.Length > 0 ? creds : null));
			    Console.WriteLine("done.");
			}
			#endregion

			#region Mac Address aquision
			Console.Write("Getting MAC addresses.... ");
			
			String macs = get_mac_addr();
			
			Console.WriteLine("done.");
			#endregion

			#region Get the machine's information
			Console.Write("Getting machine info..... ");

			String[] manufact_details = get_manufact_details();

			Console.WriteLine("done.\n");
			#endregion

			#region Disable Automatic Updates
			if (options.ContainsKey("disable_Update") && options["disable_Update"] == "true")
			{
			    disable_win_update();
			}
			#endregion

			//Get notes for the machine
			String notes = get_notes();

			#region Check for output file and set it up for CSV format
			if (!File.Exists(OUTPUT_FILE))
			{
				File.Create(OUTPUT_FILE).Close();
				StreamWriter x = new StreamWriter(OUTPUT_FILE);

				//Headers
				x.WriteLine("Name, Manufacturer, Product, S/N, MAC Addresses, Notes");

				x.Close();
			}
			#endregion

			#region Write info to output files

			//Format the output, a lot nicer to look at than string cocatenation, using \r so it avoids taking \n's from notes, macs or whatever
			String output = String.Format("\"{0}\",\"{1}\",{2},{3},\"{4}\",\"{5}\"\r", name, manufact_details[0], manufact_details[1], manufact_details[2], macs, notes);

			//Append to the output file
			File.AppendAllText(OUTPUT_FILE, output);

			#endregion

			#region Clear the console
			//Clear the console for easier reading
			try
			{
				Console.Clear();
			}catch(IOException e){}
			#endregion

			//See what the user wants to do once the program exits
			clean_up();
		}

		#region config reading
		static Dictionary<String, String> read_config()
		{
			//Return var
			Dictionary<string, string> opt_prog = new Dictionary<string, string>();

			//Opens the config file for reading
			FileStream config = File.Open(CONFIG_FILE, FileMode.Open);

			//Setup for reading from the file
			byte[] options = new byte[config.Length];
			UTF8Encoding temp = new UTF8Encoding(true);

			//Read from the file
			while (config.Read(options, 0, options.Length) > 0)
			{
				//Get the options
				string[] options_vals = temp.GetString(options).Split('\n');

				//Loop through and add the options to the list
				for (int i = 0; i < options_vals.Length; i++)
				{
				        if (options_vals[0] != "#" && options_vals[i].Length > 1)
					{
						//These temp vars are for cleanliness
						string temp1 = options_vals[i].Substring(0, options_vals[i].IndexOf(" "));
						string temp2;
						
						if (options_vals[i].Contains("\r") || options_vals[i].Contains("\n"))
						{
							temp2 = options_vals[i].Substring(options_vals[i].IndexOf(" ") + 1, options_vals[i].Length - 2 - temp1.Length);
						}
						else
						{
							temp2 = options_vals[i].Substring(options_vals[i].IndexOf(" ") + 1);
						}

	
						opt_prog.Add(temp1, temp2);
					}
				}
			}
			return opt_prog;
		}
		#endregion

		#region Build the next name for the entry
		static String get_name(String stem, String start)
		{
			//If the out file doesn't exist, just give it a default and return it
			if (!File.Exists(OUTPUT_FILE))
			{
				return stem + (start == "" ? "1" : start);
			}

			//Read from the output file if it exists
			StreamReader input = new StreamReader(OUTPUT_FILE);
			String worker = input.ReadToEnd();

			//Split it up so I can get access to the last entry name
			string[] entries = worker.Split('\r');
			string[] last_entry = entries[entries.Length - 2].Split(',');
			string current = last_entry[0].Replace("\"", "");
			String name = current.Substring(current.LastIndexOf(stem[stem.Length - 1]) + 1);

			//Get the number
			int num = Int32.Parse(name) + 1;
			input.Close();

			//Return the name
			return stem + num;
		}
		#endregion

		#region Machine Name changing code
		static void changeMachineName(String name, String[] creds)
		{
			Boolean on_ad = false;
			SelectQuery wmi = new SelectQuery("Win32_ComputerSystem");
			ManagementObjectSearcher x = new ManagementObjectSearcher(wmi);

			object[] creds_out = new object[3];

			try
			{
			    Domain.GetComputerDomain();
			    on_ad = true;
			}catch(ActiveDirectoryObjectNotFoundException){}

			foreach (ManagementObject i in x.Get())
			{                
			    creds_out[0] = name;
			    if (creds == null && on_ad)
			    {
			        creds = get_creds(null, "");
			    }
			    if (creds != null)
			    {
			        creds_out[1] = creds[2];
			        creds_out[2] = creds[1];
			    }
								   
			   i.InvokeMethod("Rename", creds_out);
			}
		}
		#endregion

		#region Gets the mac addresses of the client
		static string get_mac_addr()
		{
			NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces();
			String out_macs = "";
			bool[] done_type = new bool[2];

			foreach (NetworkInterface i in nics)
			{
				if (!i.Description.Equals("Microsoft Wi-Fi Direct Virtual Adapter") &&
					  ((i.NetworkInterfaceType == NetworkInterfaceType.Ethernet && !done_type[0]) ||
					  (i.NetworkInterfaceType == NetworkInterfaceType.Wireless80211 && !done_type[1]))) 
				{


					if (i.NetworkInterfaceType == NetworkInterfaceType.Ethernet)
					{
						done_type[0] = true;
						out_macs += "Ethernet: " + i.GetPhysicalAddress().ToString().ToLower() + "\n";
					}
					else
					{
						done_type[1] = true;
						out_macs += "WiFi: " + i.GetPhysicalAddress().ToString().ToLower() + "\n";
					}

				}
			}

			return out_macs;
		}
		#endregion

		#region Gets system info
		static String[] get_manufact_details()
		{
			String[] output = new String[3];

			SelectQuery query = new SelectQuery(@"Select * from Win32_ComputerSystem");

			using (ManagementObjectSearcher searcher2 = new ManagementObjectSearcher(query))
			{
				//execute the query
				foreach (ManagementObject process in searcher2.Get())
				{
					//print system info
					process.Get();
					
					using (ManagementObjectSearcher searcher1 = new ManagementObjectSearcher(query))
					{
						//execute the query
						foreach (ManagementObject process1 in searcher1.Get())
						{
							//print system info
							process1.Get();

							//Store it
							output[0] = process1["Manufacturer"].ToString();
							output[1] = process1["Model"].ToString();
						}
					}
				}
				searcher2.Dispose();
			}

			ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_BIOS");
			ManagementObjectCollection info = searcher.Get();

			foreach (ManagementObject i in info)
			{
				output[2] = i.GetPropertyValue("SerialNumber").ToString();
			}

			return output;
		}
		#endregion

		#region Note taking function
		static String get_notes()
		{
			String output = "";

			String temp = "";
			do
			{
				Console.Write("Enter your notes (type \"exit\" to exit) ");
				temp = Console.ReadLine();
				
				if (temp != "exit")
				{
					output += temp + "\n";
				}
			}while (temp != "exit");

			return output;
		}
		#endregion

		#region Active Directory Domain joining function
		static String[] join_AD_Domain(String domain, String uname)
		{
			//Get the credentials for the domain
			String[] creds = get_creds(domain, uname);


			using (ManagementObject wmiObject = new ManagementObject(new ManagementPath("Win32_ComputerSystem.Name='" + Environment.MachineName + "'")))
			{
				ManagementBaseObject inParams = wmiObject.GetMethodParameters("JoinDomainOrWorkgroup");

				#region Joining Credentials
				inParams["Name"] = creds[0];
				inParams["Password"] = creds[2];
				inParams["UserName"] = creds[1];
				inParams["FJoinOptions"] = 3;
				#endregion

				//Join the Domain
				ManagementBaseObject joinParams = wmiObject.InvokeMethod("JoinDomainOrWorkgroup", inParams, null);

				#region Error check
				String message = "";
				switch(joinParams["ReturnValue"].ToString()){
					case "0":
						Console.WriteLine("You have joined the domain successfully!");
						return creds;

					case "5":
						Console.WriteLine( "Access has been denied.");
						break;

					//Special Case, re-call function if error is incorrect uname/pass, less closing and reopening of program
					case "1326":
						Console.WriteLine("Incorrect Username or Password.");
						Console.WriteLine();

						join_AD_Domain(creds[0], "");
						break;

					case "1355":
						Console.WriteLine("The domain either doesn't exist or hasn't responded.");
						break;

					case "2691":
						Console.WriteLine("The machine is already joined to the domain");
						break;

					//Something weird happened
					default:
						message = "Unknown error, look up this error code " + joinParams["ReturnValue"] + " and check C:\\Windows\\debug\\NetSetup.log";
						break;
				}

				Console.Write("Press any key to continue");
				Console.ReadKey();
				#endregion
					
				return creds;
			}
		}
        #endregion

        #region Get the user credentials for the domain
        static String[] get_creds(String domain, String user)
		{
			String[] output = {domain, user, ""};

			#region Getting the domain name from console (if not provided in config file)
			while(output[0] == "" && output[0] != null)
			{
				Console.Write("Enter your Active Directory Domain name: ");
				output[0] = Console.ReadLine();
			}
			#endregion

			#region Getting the username
			while(output[1] == "")
			{
				Console.Write("Enter in the user you want to connect with: ");
				output[1] = Console.ReadLine();
			}
			#endregion

			#region Gets the password
			do
			{
				Console.Write("Enter your password: ");
				ConsoleKeyInfo key;

				#region Mask the input
				do
				{
					//Grab the console key
					key = Console.ReadKey(true);

					//Make sure it isn't backspace or enter
					if(key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
					{
						//Add the character and print a *
						output[2] += key.KeyChar;
						Console.Write("*");
					}

					//Deleting chars from the password
					else if (key.Key == ConsoleKey.Backspace && output[2].Length > 0)
					{
						output[2] = output[2].Substring(0, output[2].Length - 1);
						Console.Write("\b \b");
					}

				//Loop until Enter
				}while(key.Key != ConsoleKey.Enter);
				#endregion

			Console.WriteLine();

			//Make sure the pass isn't empty
			} while(output[2] == "");
			#endregion

			return output;
		}
        #endregion

        #region Ask the user how they want to exit the program
        static void clean_up()
		{
			Console.WriteLine("The info has been gathered!");
			String resp = "";

			do
			{
				Console.Write("Would you like to Shutdown, Restart, or Exit? (exit is default option) ");
				resp = Console.ReadLine();
			}while(resp == "");

			switch (resp)
			{
				//Shut down after 3 seconds
				case "S":
				case "s":
				case "shutdown":
					Process proc_shutdown = new Process
					{
						StartInfo = new ProcessStartInfo
						{
							FileName = "shutdown.exe",
							Arguments = "/s /t 3",
							UseShellExecute = false,
							RedirectStandardOutput = false,
							CreateNoWindow = true
						}
					};

					proc_shutdown.Start();
					break;

				//Reboot after 3 seconds
				case "R":
				case "r":
				case "restart":
					Process proc_reboot = new Process
					{
						StartInfo = new ProcessStartInfo
						{
							FileName = "shutdown.exe",
							Arguments = "/r /t 3",
							UseShellExecute = false,
							RedirectStandardOutput = false,
							CreateNoWindow = true
						}
					};

					proc_reboot.Start();
					break;

			}
		}
        #endregion

        #region Disable Windows Update (Win7 and 8)
        static void disable_win_update()
		{
			AutomaticUpdates auc = new AutomaticUpdates();

			auc.Settings.NotificationLevel = AutomaticUpdatesNotificationLevel.aunlNotifyBeforeDownload;
			auc.Settings.Save();
		}
        #endregion
    }
}
