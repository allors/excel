namespace Allors.Excel.Interop.Dna
{
    using Application;

    internal class ServiceLocator : IServiceLocator
    {
        public IAlerter Alerter { get; } = new Alerter();
    }
}
