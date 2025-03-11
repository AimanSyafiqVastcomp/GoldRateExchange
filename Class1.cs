using System.ComponentModel;
using System.Configuration.Install;
using System.ServiceProcess;

namespace GoldRatesExtractor
{
    [RunInstaller(true)]
    public class GoldRatesServiceInstaller : Installer
    {
        public GoldRatesServiceInstaller()
        {
            // Service Process Installer
            var processInstaller = new ServiceProcessInstaller
            {
                Account = ServiceAccount.LocalSystem
            };

            // Service Installer
            var serviceInstaller = new ServiceInstaller
            {
                StartType = ServiceStartMode.Automatic,
                ServiceName = "GoldRatesExtractor",
                DisplayName = "Gold Rates Extractor Service",
                Description = "Extracts gold rates from configured websites and stores them in a database."
            };

            // Add installers to collection
            Installers.Add(processInstaller);
            Installers.Add(serviceInstaller);
        }
    }
}