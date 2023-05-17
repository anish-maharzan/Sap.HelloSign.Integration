using System.ServiceProcess;

namespace Sap.HelloSign.Integration
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new HelloSignService()
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
