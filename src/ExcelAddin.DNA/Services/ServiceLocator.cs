using Application;

namespace ExcelAddIn.DNA
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