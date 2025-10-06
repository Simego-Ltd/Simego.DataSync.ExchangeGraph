using Simego.DataSync.Engine;
using Simego.DataSync.Interfaces;
using Simego.DataSync.OAuth;
using Simego.DataSync.Providers;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace Simego.DataSync.ExchangeGraph
{
    [ProviderInfo(Name = "Exchange Graph API Connector", Description = "Connect to Exchange Mailbox via Graph API", Group = "Microsoft Exchange")]
    public class ExchangeGraphDatasourceReader : DataReaderOAuth2ProviderBase, IDataSourceSetup
    {
        private const string ODATA_NEXT_LINK = "@odata.nextLink";
        private const string ODATA_COLUMNS = "id,internetMessageId,subject,receivedDateTime";

        private ConnectionInterface _connectionIf;
        private OAuthWebConnectionWrapper _oAuthWebConnectionWrapper = new OAuthWebConnectionWrapper();

        private string _accessToken;
        private DateTime _tokenExpires;

        [Category("Authentication")]
        public string TenantId { get; set; }

        [Category("Authentication")]
        public string ClientId { get; set; }

        [Category("Settings")]
        [Description("User Mailbox to read mail from.")]
        public string UserPrincipalName { get; set; }

        [Category("Settings")]
        [Description("Email address of messages to return from Mailbox.")]
        public string SenderEmail { get; set; }

        internal string ClientSecret { get; set; }
        
        public override DataTableStore GetDataTable(DataTableStore dt)
        {
            dt.AddIdentifierColumn(typeof(string));

            var mapping = new DataSchemaMapping(SchemaMap, Side);
            var columns = SchemaMap.GetIncludedColumns();
            
            // Setup Web Request Helper
            var helper = GetRequestHelper();
           
            var url = $"https://graph.microsoft.com/v1.0/users/{UserPrincipalName}/messages?$filter=from/emailAddress/address eq '{SenderEmail}'&$select={ODATA_COLUMNS}";

            do
            {
                var response = helper.GetRequestAsJson(url);
                
                if (response["value"] != null)
                {
                    foreach (var item_row in response["value"])
                    {
                        if (dt.Rows.AddWithIdentifier(mapping, columns,
                                (item, columnName) =>
                                {
                                    return item_row[columnName] == null ? null : item_row[columnName].ToObject<object>();
                                }, item_row["id"].ToObject<string>()) == DataTableStore.ABORT)
                        {
                            return dt;
                        }
                    }
                }

                url = response[ODATA_NEXT_LINK]?.ToObject<string>();

            } while (url != null);

            return dt;
        }

        public override DataSchema GetDefaultDataSchema()
        {
            // Return the Data source default Schema.

            var schema = new DataSchema();

            schema.Map.Add(new DataSchemaItem("id", typeof(string), true, false, false, -1));
            schema.Map.Add(new DataSchemaItem("internetMessageId", typeof(string), false, false, false, -1));
            schema.Map.Add(new DataSchemaItem("subject", typeof(string), false, false, true, -1));
            schema.Map.Add(new DataSchemaItem("receivedDateTime", typeof(DateTime), false, false, false, -1));

            return schema;

        }

        public override List<ProviderParameter> GetInitializationParameters()
        {
            //Return the Provider Settings so we can save the Project File.
            return new List<ProviderParameter>
                       {
                            new ProviderParameter(nameof(TenantId), TenantId),
                            new ProviderParameter(nameof(ClientId), ClientId),
                            new ProviderParameter(nameof(ClientSecret), SecurityService.EncryptValue(ClientSecret)),
                            new ProviderParameter(nameof(UserPrincipalName), UserPrincipalName),
                            new ProviderParameter(nameof(SenderEmail), SenderEmail)
                       };
        }

        public override void Initialize(List<ProviderParameter> parameters)
        {
            //Load the Provider Settings from the Project File.
            foreach (ProviderParameter p in parameters)
            {
                switch (p.Name)
                {
                    case nameof(TenantId):
                        {
                            TenantId = p.Value;
                            break;
                        }
                    case nameof(ClientId):
                        {
                            ClientId = p.Value;
                            break;
                        }
                    case nameof(ClientSecret):
                        {
                            ClientSecret = SecurityService.DecyptValue(p.Value);
                            break;
                        }
                    case nameof(UserPrincipalName):
                        {
                            UserPrincipalName = p.Value;
                            break;
                        }
                    case nameof(SenderEmail):
                        {
                            SenderEmail = p.Value;
                            break;
                        }
                }
            }
        }

        public override IDataSourceWriter GetWriter() => new NullWriterDataSourceProvider { SchemaMap = SchemaMap };
        
        public void DisplayConfigurationUI(IntPtr parent)
        {
            var parentControl = Control.FromHandle(parent);

            if (_connectionIf == null)
            {
                _connectionIf = new ConnectionInterface();
                _connectionIf.PropertyGrid.SelectedObject = new ConnectionProperties(this);
            }

            _connectionIf.Font = parentControl.Font;
            _connectionIf.Size = new Size(parentControl.Width, parentControl.Height);
            _connectionIf.Location = new Point(0, 0);
            _connectionIf.Dock = DockStyle.Fill;

            parentControl.Controls.Add(_connectionIf);
        }

        public bool Validate() => true;

        public IDataSourceReader GetReader() => this;
                
        public override string GetFileName(DataCompareItem item, int index) => $"{item.GetSourceIdentifier<string>()}.eml";

        public override string GetFilePath(DataCompareItem item, int index) => string.Empty;

        public override string GetBlobTempFile(DataCompareItem item, int index)
        {
            var helper = GetRequestHelper();
            var id = item.GetSourceIdentifier<string>();
            var fileName = FileCache.GetTempFileName();
            
            using (var fs = File.Create(fileName))
            {
                var stream = helper.OpenReadStream($"https://graph.microsoft.com/v1.0/users/{UserPrincipalName}/messages/{id}/$value");

                stream.CopyTo(fs);
            }

            return fileName;
        }

        public override OAuthConfiguration GetOAuthConfiguration()
        {
            return new OAuthConfiguration
            {
                TokenUrl = $"https://login.microsoftonline.com/{TenantId}/oauth2/v2.0/token",
                ClientID = ClientId,
                ClientSecret = ClientSecret,
                GrantType = "client_credentials",
                Scope = "https://graph.microsoft.com/.default",

                AccessToken = _accessToken,
                TokenExpires = _tokenExpires
            };
        }

        public override void UpdateOAuthConfiguration(OAuthConfiguration configuration)
        {
            _accessToken = configuration.AccessToken;
            _tokenExpires = configuration.TokenExpires;
        }

        private HttpWebRequestHelper GetRequestHelper()
        {
            // Get an Access Token
            BeginOAuthAuthorize(_oAuthWebConnectionWrapper);

            var helper = new HttpWebRequestHelper();
            helper.SetAuthorizationHeader(_accessToken);
            return helper;
        }

    }
}
