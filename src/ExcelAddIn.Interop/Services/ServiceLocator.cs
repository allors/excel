using Application;

namespace ExcelAddInLocal
{
    internal class ServiceLocator : IServiceLocator
    {
        public ServiceLocator()
        {
            this.Alerter = new Alerter();
        }

        public IAlerter Alerter { get; }
    }
}