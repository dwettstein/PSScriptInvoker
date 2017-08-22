using System.ServiceProcess;

namespace PSScriptInvoker
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
                new PSScriptInvoker()
            };
            ServiceBase.Run(ServicesToRun);
            // The following code is just used for debugging in Visual Studio.
            //(new PSScriptInvoker()).StartPSScriptInvoker().Wait();
        }
    }
}
