using Application;

namespace ExcelAddIn.VSTO
{
    internal class ServiceLocator : IServiceLocator
    {
        public IAlerter Alerter { get; } = new Alerter();
    }
}
