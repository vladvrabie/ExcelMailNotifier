using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.Threading.Tasks;

namespace MailingWindowsService
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : System.Configuration.Install.Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();

            //serviceInstaller1.ServiceName = "...";
            //serviceInstaller1.DisplayName = "...";
            //serviceInstaller1.Description = "...";
        }
    }
}
