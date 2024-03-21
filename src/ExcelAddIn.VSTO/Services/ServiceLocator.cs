using Application;

namespace ExcelAddIn.VSTO
{
    internal class ServiceLocator : IServiceLocator
    {
        public ServiceLocator() => this.Alerter = new Alerter();

        public IAlerter Alerter { get; }
    }
}
