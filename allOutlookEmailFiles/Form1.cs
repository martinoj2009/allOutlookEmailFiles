using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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
            //Get all the SIDS in the registry
            string[] users = Registry.Users.GetSubKeyNames();

            foreach(string user in users)
            {
                //As long as the string isn't the DEFAULT user then continue
                if(!(user == ".DEFAULT"))
                {
                    registryUsers.Items.Add(user);
                }
                
            }

            

        }


        //This is the background worker that will get the files
        private void GETFILES_DoWork(object sender, DoWorkEventArgs e)
        {

        }

        //Perform this task when a SID is selected
        private void registryUsers_SelectedIndexChanged(object sender, EventArgs e)
        {
            RegistryKey usersRegistry = Registry.Users.OpenSubKey(registryUsers.SelectedItem.ToString() + @"\SOFTWARE\Microsoft\Office\16.0\Outlook\Search");

            string[] outlookFiles = null;
            try
            {
                outlookFiles = usersRegistry.GetValueNames();
            }
            catch(Exception)
            {

            }
            
            if(outlookFiles != null)
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
            string[] selectedItems;

            foreach(string item in allOutlookFiles.SelectedItems)
            {

            }

        }
    }
}
