using Application;

namespace ExcelAddIn.Services
{
    internal class ServiceLocator : IServiceLocator
    {
        public ServiceLocator()
        {
            Alerter = new Alerter();
        }

        public IAlerter Alerter { get; }
    }
}