using System.Collections.Generic;
using System.Linq;
using LinqToDB.Configuration;

namespace Watcher
{
    public class LinqToDBSettings : ILinqToDBSettings
    {
        public IEnumerable<IDataProviderSettings> DataProviders => Enumerable.Empty<IDataProviderSettings>();

        public string DefaultConfiguration => "SqlServer";
        public string DefaultDataProvider => "SqlServer";

        public IEnumerable<IConnectionStringSettings> ConnectionStrings
        {
            get
            {
                yield return new ConnectionStringSettings {
                    Name = "SharpLayout",
                    ProviderName = "SqlServer",
                    ConnectionString = @"Data Source=(local)\SQL2014;Initial Catalog=SharpLayout;Integrated Security=True"
                };
            }
        }
    }
    
    public class ConnectionStringSettings : IConnectionStringSettings
    {
        public string ConnectionString { get; set; }
        public string Name { get; set; }
        public string ProviderName { get; set; }
        public bool IsGlobal => false;
    }
}