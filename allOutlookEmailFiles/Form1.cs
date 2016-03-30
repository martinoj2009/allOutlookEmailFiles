using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Management;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace allOutlookEmailFiles
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            main();
        }


        private void main()
        {
            
            GETFILES.WorkerSupportsCancellation = true;
            usernameLabel.Text = System.Environment.GetEnvironmentVariable("USERNAME");
            //ManagementObjectSearcher query = new ManagementObjectSearcher("SELECT * FROM Win32_UserProfile WHERE Loaded = True");

            officeVersionLabel.Text = "";

            //Add all the SIDs on the machine, except for DEFAULT to the listbox
            foreach (string user in Directory.GetDirectories(@"C:\Users\"))
            {
                //As long as the string isn't the DEFAULT user then continue
                if (!(user == ".DEFAULT"))
                {
                    registryUsers.Items.Add(user.Split('\\')[2]);
                }

            }



        }


        //This is the background worker that will get the files
        private void GETFILES_DoWork(object sender, DoWorkEventArgs e)
        {
            //This is optional stuff the computer can get
            //Get the username
            
        }

        //Perform this task when a SID is selected
        private void registryUsers_SelectedIndexChanged(object sender, EventArgs e)
        {
            officeVersionLabel.Text = getOutlookVersion();
            RegistryKey usersRegistry = Registry.Users.OpenSubKey(registryUsers.SelectedItem.ToString() + @"\SOFTWARE\Microsoft\Office\" + getOutlookVersion() + @"\Outlook\Search");

            if (registryUsers.SelectedItem == null)
            {
                return;
            }
            try
            {
                string account = new System.Security.Principal.SecurityIdentifier(registryUsers.SelectedItem.ToString()).Translate(typeof(System.Security.Principal.NTAccount)).ToString();
                usernameLabel.Text = account;

            }
            catch (Exception)
            {
                usernameLabel.Text = "";
            }


            ArrayList outlookFiles = getOutlookFiles();

            if(outlookFiles == null)
            {
                allOutlookFiles.Items.Clear();
                allOutlookFiles.Items.Add("No Files Found");
                return;
            }
            /*
            try
            {
                outlookFiles = usersRegistry.GetValueNames();
            }
            catch(Exception)
            {

            }
            */
            if(outlookFiles != null || outlookFiles.Count != 0)
            {
                allOutlookFiles.Items.Clear();
                foreach (string name in outlookFiles)
                {
                    allOutlookFiles.Items.Add(name);
                }
            }
            else
            {
                allOutlookFiles.Items.Clear();
                allOutlookFiles.Items.Add("NO FILES FOUND!");
            }

            

            
        }

        private void allOutlookFiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Do nothing if nothing is selected
            if(allOutlookFiles.SelectedItems.Count == 0)
            {
                return;
            }

            string items = "";
            
            foreach(string item in allOutlookFiles.SelectedItems)
            {
                items = items + item + Environment.NewLine;
            }

            Clipboard.SetText(items);
        }

        private string[] getSIDS()
        {
            //Get all the SIDS in the registry
            string[] users = Registry.Users.GetSubKeyNames();
            return users;
        }


        private string getOutlookVersion()
        {
            string highestVersion = "0";

            
            //This function will get the registry of the requested user
            Process p = new Process();
            ProcessStartInfo info = new ProcessStartInfo();
            info.FileName = "getVersion.bat";
            info.Arguments = "ITSUS";
            info.RedirectStandardInput = true;
            info.UseShellExecute = false;
            info.RedirectStandardOutput = true;
            info.CreateNoWindow = true;
            info.WindowStyle = ProcessWindowStyle.Hidden;

            p.StartInfo = info;
            p.Start();
            string[] output = p.StandardOutput.ReadToEnd().Split('\n');

            ArrayList officeVersions = new ArrayList();
            foreach (string line in output)
            {
                if(line.Contains("ERROR:"))
                {
                    return "";
                }

                
                if(line.StartsWith(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\"))
                {
                    try
                    {
                        Double.Parse(line.Split('\\')[4]);
                        officeVersions.Add(line.Split('\\')[4]);
                    }
                    catch(Exception)
                    {
                        //Do nothing as it's just a string
                    }
                    
                }
            }

            foreach (string num in officeVersions)
            {
                if (Double.Parse(num) > Double.Parse(highestVersion))
                {
                    highestVersion = num;
                }
            }


            return highestVersion;
        }

        private ArrayList getOutlookFiles()
        {
            //This function will get the registry of the requested user
            Process p = new Process();
            ProcessStartInfo info = new ProcessStartInfo();
            info.FileName = "getFiles.bat";
            info.Arguments = registryUsers.SelectedItem + " " + officeVersionLabel.Text;
            info.RedirectStandardInput = true;
            info.UseShellExecute = false;
            info.RedirectStandardOutput = true;
            info.CreateNoWindow = true;
            info.WindowStyle = ProcessWindowStyle.Hidden;

            p.StartInfo = info;
            string[] output = null;
            try
            {
                p.Start();
                output = p.StandardOutput.ReadToEnd().Split('\n');
            }
            catch(Exception)
            {

            }

            ArrayList files = new ArrayList();
            bool filesFound = false;

            if(output != null)
            {
                foreach (string line in output)
                {
                    //files.Add(line);
                    if (line.Contains("ERROR:"))
                    {

                        return files = (null);
                    }

                    if (line.Contains("REG_DWORD"))
                    {
                        filesFound = true;
                        //MessageBox.Show(line.Split(new string[] { "REG_DWORD" }, StringSplitOptions.None)[0]);
                        files.Add(line.Split(new string[] { "REG_DWORD" }, StringSplitOptions.None)[0].Trim());
                    }
                }

                if(filesFound == false)
                {
                    return files = (null);
                }
            }
            


            return files;
        }

    }
}
