using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.Threading.Tasks;
using DCRM_Utils;
using Microsoft.Crm.Sdk.Messages;
using System.Threading;

namespace BatchImportIncidents
{
    class Demande
    {
        #region CONST Strings
        private const string HEADER_ACCOUNT = "Compte";
        private const string HEADER_CLIENT = "Client";
        private const string HEADER_DESCRIPTION = "Description";
        private const string HEADER_THEME1 = "Thème 1";
        private const string HEADER_THEME2 = "Thème 2";
        private const string HEADER_THEME3 = "Thème 3";
        private const string HEADER_TITRE = "Titre";
        private const string HEADER_STATUT = "Statut";
        private const string HEADER_RAISON_STATUT = "Raison du statut";
        private const string HEADER_GUID_DEMANDE = "Guid Demande";

        private const string GUID_THEME1 = "THEME1_GUID";
        private const string GUID_THEME2 = "THEME2_GUID";
        private const string GUID_THEME3 = "THEME3_GUID";

        private const string VAUE_STATUT_RESOLVED = "Résolu";
        #endregion // CONST Strings

        #region Attributes

        public delegate void IncidentCreator(int currentIndex, int maxIndex, bool isCallingResolve);
        public delegate void IncidentResolver(int currentIndex, int maxIndex, Guid incidentId);

        private readonly Dictionary<string, string> AccountDict = new Dictionary<string, string>();
        private Dictionary<string, string> IncidentReadDict = new Dictionary<string, string>
        {
            { HEADER_GUID_DEMANDE, ""},
            { HEADER_DESCRIPTION, ""},
            { HEADER_THEME1, ""},
            { HEADER_THEME2, ""},
            { HEADER_THEME3, ""},
            { HEADER_TITRE, ""},
            { HEADER_ACCOUNT, ""},
            { HEADER_STATUT, ""},
            { HEADER_RAISON_STATUT, ""}
        };
        private Dictionary<string, string> IncidentWriteDict = new Dictionary<string, string>
        {
            { HEADER_DESCRIPTION, ""},
            { GUID_THEME1, ""},
            { GUID_THEME2, ""},
            { GUID_THEME3, ""},
            { HEADER_TITRE, ""},
            { HEADER_CLIENT, ""},
            { HEADER_STATUT, ""},
            { HEADER_RAISON_STATUT, ""}
        };
        #endregion // Attributes

        #region SetFieldValue
        public void SetFieldValue(string fieldKey, string filedValue)
        {
            if (IncidentReadDict.Keys.Contains(fieldKey))
                IncidentReadDict[fieldKey] = filedValue;
        }
        #endregion // SetFieldValue

        #region GetEntityValue
        private string GetEntityValue(Entity entity, string entityKey)
        {
            var value = entity[entityKey];
            var valueString = "";

            if (value.GetType() == typeof(string))
                valueString = (string)value;

            if (value.GetType() == typeof(Guid))
                valueString = value.ToString();

            if (value.GetType() == typeof(EntityReference))
            {
                EntityReference reference = (EntityReference)value;
                valueString = reference.Id.ToString();
            }

            return (valueString);
        }
        #endregion // GetEntityValue

        #region GetGuidFromEntity
        private string GetGuidFromEntity(string entityName, string iDLookupKey, string entityLookupKey, string entityLookupValue)
        {
            var ctx = DcrmConnectorFactory.GetContext();

            var entityQuery = from entity in ctx.CreateQuery(entityName)
                              where entity[entityLookupKey].Equals(entityLookupValue)
                              select new
                              {
                                  Guid = GetEntityValue(entity, iDLookupKey)
                              };

            var guid = "";
            try
            {
                var entity = entityQuery.First();
                if (entity != null)
                {
                    guid = entity.Guid.ToString();
                }
            }
            catch (Exception Ex)
            {
                MiscHelper.WriteLine($"GetGuidFromEntity : {Ex.Message}");
            }

            return guid;
        }
        #endregion // GetGuidFromAccount

        #region Resolve
        public Thread Resolve(int currentIndex, int maxIndex, Guid incidentId, bool isPrintingProgress)
        {
            Thread th = null;

            var ctx = DcrmConnectorFactory.GetContext();
            var dcrmConnector = DcrmConnectorFactory.Get();
            var srv = dcrmConnector.GetService();
            var incident = new EntityReference(Incident.EntityLogicalName, incidentId);

            IncidentResolution incidentResolution = new IncidentResolution
            {
                IncidentId = incident,
                StatusCode = new OptionSetValue(5)
            };

            CloseIncidentRequest closeIncidentReq = new CloseIncidentRequest()
            {
                IncidentResolution = incidentResolution,
                Status = new OptionSetValue(5)
            };

            try
            {
                th = new Thread(() =>
                {
                    ctx.Execute(closeIncidentReq);
                    if (isPrintingProgress)
                        MiscHelper.IncrementProgressBar(maxIndex);
                });
                th.Start();
            }
            catch (Exception ex)
            {
                var message = string.Format($"Could not close the incident : {incidentId} : {ex.Message} ");
                MiscHelper.WriteLine(message);
            }
            return th;
        }
        #endregion // Resolve

        #region Create
        public Thread Create(int currentIndex, int maxIndex, bool isCallingResolve)
        {
            Thread th = new Thread(() =>
            {
                IncidentCreator Create = CreateAndCloseIncident;
                Create(currentIndex, maxIndex, isCallingResolve);
            });
            th.Start();

            return th;
        }
        #endregion // Create

        #region CreateAndCloseIncident
        public void CreateAndCloseIncident(int currentIndex, int maxIndex, bool isCallingResolve)
        {
            var ctx = DcrmConnectorFactory.GetContext();
            var dcrmConnector = DcrmConnectorFactory.Get();
            var srv = dcrmConnector.GetService();
            Guid incidentId = Guid.Empty;

           var createIncidentTimer = new MiscHelper();
            createIncidentTimer.StartTimeWatch();

            var accountId = this.IncidentWriteDict[HEADER_CLIENT];
            if (string.IsNullOrEmpty(accountId))
            {
                var accountNumber = this.IncidentWriteDict[HEADER_ACCOUNT];
                MiscHelper.WriteLine($"GUID for account : {accountNumber} won't be created");
                return;
            }

            var targetIncident = new Entity("incident");
            targetIncident["customerid"] = new EntityReference(Account.EntityLogicalName, new Guid(accountId));
            targetIncident["title"] = this.IncidentWriteDict[HEADER_TITRE];
            targetIncident["description"] = this.IncidentWriteDict[HEADER_DESCRIPTION];

            var guidTheme1 = new Guid(this.IncidentWriteDict[GUID_THEME1]);
            var theme1Entity = new EntityReference("crm_themeniveau1", guidTheme1);
            targetIncident["crm_theme_niveau_1_id"] = theme1Entity;

            var guidTheme2 = new Guid(this.IncidentWriteDict[GUID_THEME2]);
            var theme2Entity = new EntityReference("crm_themeniveau2", guidTheme2);
            targetIncident["crm_theme_niveau_2_id"] = theme2Entity;

            var guidTheme3 = new Guid(this.IncidentWriteDict[GUID_THEME3]);
            var theme3Entity = new EntityReference("crm_themeniveau3", guidTheme3);
            targetIncident["crm_theme_niveau_3_id"] = theme3Entity;

            incidentId = srv.Create(targetIncident);
            if (isCallingResolve)
                this.Resolve(currentIndex, maxIndex, incidentId, false);
            createIncidentTimer.StopTimeWatch();
            
            MiscHelper.IncrementProgressBar(maxIndex);
        }
        #endregion // CreateAndCloseIncident

        #region Process
        public void Process()
        {
            var sb = new StringBuilder();

            foreach (var headerKey in IncidentReadDict.Keys)
            {
                switch (headerKey)
                {
                    case HEADER_ACCOUNT:
                        {
                            var accountNumber = IncidentReadDict[headerKey];
                            var guid = GetGuidFromEntity("account", "accountid", "accountnumber", accountNumber);
                            if (string.IsNullOrEmpty(guid))
                            {
                                MiscHelper.WriteLine($"GUID for account : {accountNumber} could not be found in DCRM");
                            }
                            else
                            {
                                IncidentWriteDict[HEADER_CLIENT] = guid;
                            }
                        }
                        break;

                    case HEADER_THEME3:
                        {
                            var theme3Name = IncidentReadDict[HEADER_THEME3];
                            var theme3Id = GetGuidFromEntity("crm_themeniveau3", "crm_themeniveau3id", "crm_name", theme3Name);
                            var theme2Id = GetGuidFromEntity("crm_themeniveau3", "crm_theme_niveau_2_id", "crm_themeniveau3id", theme3Id);
                            var theme1Id = GetGuidFromEntity("crm_themeniveau2", "crm_theme_niveau_1_id", "crm_themeniveau2id", theme2Id);

                            IncidentWriteDict[GUID_THEME1] = theme1Id;
                            IncidentWriteDict[GUID_THEME2] = theme2Id;
                            IncidentWriteDict[GUID_THEME3] = theme3Id;
                        }
                        break;

                    default:
                        IncidentWriteDict[headerKey] = IncidentReadDict[headerKey];
                        break;
                }
            }
        }
        #endregion // Process

        #region GetGuid
        public Guid GetGuid()
        {
            return new Guid(IncidentReadDict[HEADER_GUID_DEMANDE]);
        }
        #endregion // GetGuid
    }
}
