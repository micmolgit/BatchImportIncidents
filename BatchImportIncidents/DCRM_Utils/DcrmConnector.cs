using System;
using Microsoft.Crm.Sdk.ServiceHelper;
using Microsoft.Xrm.Sdk.Client;

namespace DCRM_Utils
{
    class DcrmConnector
    {
        #region Properties
        private ServerConnection _serverConnection;
        private ServerConnection.Configuration _serverConfig;
        private OrganizationServiceProxy _serviceProxy;
        private ServiceContext _serviceContext;
        public ServiceContext SrvContext
        {
            get
            {
                if (_serviceContext == null)
                {
                    if (_serviceProxy == null)
                        this.Connect();
                    _serviceContext = new ServiceContext(_serviceProxy);
                    // This statement is required to enable early-bound type support.
                    _serviceProxy.EnableProxyTypes();
                }

                if (_serviceContext == null)
                {
                    throw new ArgumentNullException("ServiceContext", "ServiceContext Could not be created from OrganizationServiceProxy");
                }
                return _serviceContext;
            }
        }
        #endregion // Properties

        #region Constructor
        public DcrmConnector()
        {
            Connect();
        }
        #endregion // Constructor

        #region Connect
        public void Connect()
        {
            if (_serverConnection == null)
            {
                // Obtain the target organization's Web address and client logon
                // credentials from the user.
                _serverConnection = new ServerConnection();
                _serverConfig = _serverConnection.GetServerConfiguration();

                _serviceProxy = new OrganizationServiceProxy(_serverConfig.OrganizationUri, _serverConfig.HomeRealmUri, _serverConfig.Credentials, _serverConfig.DeviceCredentials);

                _serviceProxy.Authenticate();

                // This statement is required to enable early-bound type support.
                _serviceProxy.EnableProxyTypes();

                if (!_serviceProxy.IsAuthenticated)
                {
                    throw new InvalidOperationException("Authentication could not be completed");
                }
                MiscHelper.PrintMessage($"Successfully connected to :\n{_serverConfig.OrganizationUri}\n");
            }
        }
        #endregion // Connect

        #region Disconnect
        public void Disconnect()
        {
            _serviceProxy?.Dispose();
        }
        #endregion // Disconnect

        #region GetService
        public OrganizationServiceProxy GetService()
        {
            return _serviceProxy;
        }
        #endregion // GetService
    }

    class DcrmConnectorFactory
    {
        #region Properties
        private static DcrmConnector _dcrmConnector;
        #endregion // Properties

        #region Create
        private static DcrmConnector Create()
        {
            _dcrmConnector = new DcrmConnector();
            return _dcrmConnector;
        }
        #endregion // Create

        #region Get
        public static DcrmConnector Get()
        {
            if (_dcrmConnector == null)
                _dcrmConnector = DcrmConnectorFactory.Create();

            return _dcrmConnector;
        }
        #endregion // Get

        #region GetContext
        public static ServiceContext GetContext()
        {
            _dcrmConnector = DcrmConnectorFactory.Get();
            return _dcrmConnector.SrvContext;
        }
        #endregion // GetContext

        #region Close
        public static void Close()
        {
            var dcrmConnector = DcrmConnectorFactory.Get();
            if (dcrmConnector != null)
                dcrmConnector.Disconnect();
        }
        #endregion // Close
    }
}
