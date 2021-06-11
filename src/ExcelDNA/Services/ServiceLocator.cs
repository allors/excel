namespace ExcelDNA
{
    using Application;

    internal class ServiceLocator : IServiceLocator
    {
        public ServiceLocator() => this.Alerter = new Alerter();

        public IAlerter Alerter { get; }
    }
}
