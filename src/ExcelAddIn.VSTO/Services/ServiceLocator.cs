using Application;

namespace ExcelAddIn.VSTO
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