using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace allOutlookEmailFiles
{
    public class userAccount
    {

        public string SID { get; set; }
        public string username { get; set; }
        public string[] outlookFiles { get; set; }


        public userAccount(string setName)
        {
            username = setName;
        }

        public void test()
        {

        }
    }
}
