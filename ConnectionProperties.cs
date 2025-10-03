using System.ComponentModel;

namespace Simego.DataSync.ExchangeGraph
{
    class ConnectionProperties
    {
        private readonly ExchangeGraphDatasourceReader _reader;

        [Category("Authentication")]
        public string TenantId { get => _reader.TenantId; set => _reader.TenantId = value; }
        
        [Category("Authentication")]
        public string ClientId { get => _reader.ClientId; set => _reader.ClientId = value; }
        
        [Category("Authentication")]
        public string ClientSecret { get => _reader.ClientSecret; set => _reader.ClientSecret = value; }

        [Category("Settings")]
        [Description("User Mailbox to read mail from.")]
        public string UserPrincipalName { get => _reader.ClientSecret; set => _reader.ClientSecret = value; }

        [Category("Settings")]
        [Description("Email address of messages to return from Mailbox.")]
        public string SenderEmail { get => _reader.ClientSecret; set => _reader.ClientSecret = value; }

        public ConnectionProperties(ExchangeGraphDatasourceReader reader)
        {
            _reader = reader;
        }        
    }
}
