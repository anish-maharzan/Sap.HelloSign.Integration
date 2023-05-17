using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace Sap.HelloSign.Integration
{
    public partial class HelloSignService : ServiceBase
    {
        public HelloSignService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Debugger.Launch();
            new HelloSign();
        }

        protected override void OnStop()
        {
        }
    }
}
