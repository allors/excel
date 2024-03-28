using Application;

namespace Allors.Excel.Interop.Vsto
{
    internal class ServiceLocator : IServiceLocator
    {
        public IAlerter Alerter { get; } = new Alerter();
    }
}
