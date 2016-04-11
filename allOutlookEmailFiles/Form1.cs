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

            officeVersionLabel.Text = "";


            //Add all the SIDs on the machine, except for DEFAULT to the listbox
            foreach (string user in Directory.GetDirectories(@"C:\Users\"))
            {
                //As long as the string isn't the DEFAULT user then continue
                if (!(user == "All Users") || !(user == "Default"))
                {
                    registryUsers.Items.Add(user.Split('\\')[2]);
                }

            }

        }


        //This is the background worker that will get the files
        private void GETFILES_DoWork(object sender, DoWorkEventArgs e)
        {
            //I'll make this later
            
        }

        //Perform this task when a SID is selected
        private void registryUsers_SelectedIndexChanged(object sender, EventArgs e)
        {
            officeVersionLabel.Text = getOutlookVersion();

            //Test if file is locked
            if(IsFileLocked(@"C:\Users\" + registryUsers.SelectedItem.ToString() + @"\NTUSER.DAT"))
            {
                MessageBox.Show("This user's NTUSER.DAT file is locked. You may need to reboot as Windows locks this file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


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


            GC.Collect();
            
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


        private string getOutlookVersion()
        {

            //Mount NTUSER
            Process mountNTUser = new Process();
            mountNTUser.StartInfo.FileName = "C:\\Windows\\System32\\reg.exe";
            mountNTUser.StartInfo.Arguments = "load HKU\\Subkey \"C:\\Users\\" + registryUsers.SelectedItem + "\\NTUSER.DAT\"";
            mountNTUser.StartInfo.CreateNoWindow = true;
            mountNTUser.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

            try
            {
                mountNTUser.Start();
                mountNTUser.WaitForExit();
            }
            catch(Exception)
            {
                return "";
            }

            //Query the key

            string highestVersion = "0";

            //This function will get the registry of the requested user
            Process p = new Process();
            ProcessStartInfo info = new ProcessStartInfo();
            info.FileName = "C:\\Windows\\System32\\reg.exe";
            info.Arguments = "query HKU\\Subkey\\SOFTWARE\\Microsoft\\Office";
            info.RedirectStandardInput = true;
            info.UseShellExecute = false;
            info.RedirectStandardOutput = true;
            info.CreateNoWindow = true;
            info.WindowStyle = ProcessWindowStyle.Hidden;

            p.StartInfo = info;
            string[] output;

            try
            {
                p.Start();
                output = p.StandardOutput.ReadToEnd().Split('\n');
                p.WaitForExit();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                p.Dispose();
                output = null;
                return "0";

            }


            ArrayList officeVersions = new ArrayList();
            foreach (string line in output)
            {
                if(line.Contains("ERROR:"))
                {
                    return "";
                }

                
                if(line.StartsWith(@"HKEY_USERS\Subkey\SOFTWARE\Microsoft\Office\"))
                {
                    //MessageBox.Show(line);
                    try
                    {
                        Double.Parse(line.Split('\\')[5]);
                        officeVersions.Add(line.Split('\\')[5]);
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


            //Mount NTUSER
            Process mountNTUser = new Process();
            mountNTUser.StartInfo.FileName = "C:\\Windows\\System32\\reg.exe";
            mountNTUser.StartInfo.Arguments = "load HKU\\Subkey \"C:\\Users\\" + registryUsers.SelectedItem + "\\NTUSER.DAT\"";
            mountNTUser.StartInfo.CreateNoWindow = true;
            mountNTUser.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;


            try
            {
                mountNTUser.Start();
                mountNTUser.WaitForExit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot mount the NTUSER.dat file. " + ex.Message);
                return null;
            }


            //This function will get the registry of the requested user
            Process p = new Process();
            ProcessStartInfo info = new ProcessStartInfo();
            info.FileName = "C:\\Windows\\System32\\reg.exe";
            info.Arguments = "query HKU\\Subkey\\SOFTWARE\\Microsoft\\Office\\" + officeVersionLabel.Text.Trim() + "\\Outlook\\Search";
            info.RedirectStandardInput = true;
            info.UseShellExecute = false;
            info.RedirectStandardOutput = true;
            info.CreateNoWindow = true;
            info.WindowStyle = ProcessWindowStyle.Hidden;


            p.StartInfo = info;
            string[] output;

            try
            {
                p.Start();
                output = p.StandardOutput.ReadToEnd().Split('\n');
                p.WaitForExit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                p.Dispose();
                output = null;
                return null;

            }


            //Unload the registry
            Process unLoad = new Process();
            unLoad.StartInfo.FileName = "C:\\Windows\\System32\\reg.exe";
            unLoad.StartInfo.Arguments = "unload HKU\\Subkey";
            unLoad.StartInfo.CreateNoWindow = true;
            unLoad.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;


            try
            {
                unLoad.Start();
                unLoad.WaitForExit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cant unload the NTUSER.dat. " + ex.Message);
                return null;
            }


            ArrayList files = new ArrayList();
            bool filesFound = false;


                foreach (string line in output)
                {
                allOutlookFiles.Items.Add(line);
                    if (line.Contains("ERROR:"))
                    {
                        p.Dispose();
                        output = null;
                        return files = (null);
                    }

                    if (line.Contains("REG_DWORD"))
                    {
                    
                        filesFound = true;
                        files.Add(line.Split(new string[] { "REG_DWORD" }, StringSplitOptions.None)[0].Trim());
                    }
                }

            p.Dispose();
            output = null;
            return files;
        }

        public bool IsFileLocked(string filePath)
        {
            try
            {
                using (File.Open(filePath, FileMode.Open)) { }
            }
            catch (IOException e)
            {
                var errorCode = System.Runtime.InteropServices.Marshal.GetHRForException(e) & ((1 << 16) - 1);

                return errorCode == 32 || errorCode == 33;
            }

            return false;
        }
    }
}
