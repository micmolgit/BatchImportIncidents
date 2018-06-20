using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using DCRM_Utils;
using System.Threading.Tasks;

namespace BatchImportIncidents
{
    class Demande
    {
        #region CONST Strings
        private const string HEADER_PROPRIETAIRE = "Propriétaire";
        private const string HEADER_OWNER = "Propriétaire";
        private const string HEADER_ACCOUNT = "Compte";
        private const string HEADER_CUSTCODE = "CustCode";
        private const string HEADER_ACCOUNT_ID = "AccoountId";
        private const string HEADER_CUSTOMER_ID = "CustomerId";
        private const string HEADER_CONTRAT = "Contrat";
        private const string HEADER_MSISDN = "Msisdn";

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

        public delegate void IncidentCreator(int currentIndex, int maxCount, bool isCallingResolve);
        public delegate void IncidentResolver(int currentIndex, int maxCount, Guid incidentId);

        private readonly Dictionary<string, string> AccountDict = new Dictionary<string, string>();
        private Dictionary<string, string> IncidentReadDict = new Dictionary<string, string>();       
        private Dictionary<string, string> IncidentWriteDict = new Dictionary<string, string>();       

        public bool IsCallingResolve { get; set; }
        #endregion // Attributes

        #region SetFieldValue
        public void SetFieldValue(string fieldKey, string filedValue)
        {
            //if (IncidentReadDict.Keys.Contains(fieldKey))
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
                MiscHelper.PrintMessage($"GetGuidFromEntity : {Ex.Message}");
            }

            return guid;
        }
        #endregion // GetGuidFromAccount

        #region GetCrmClientFormCustCode
        private CrmClient GetCrmClientFormCustCode(string custCode)
        {
            CrmClient client = null;

            var ctx = DcrmConnectorFactory.GetContext();

            var partyQuery = from party in ctx.CreateQuery("account")
                             join client_facturation in ctx.CreateQuery("crm_clientdefacturation")                             
                              on party["accountid"] equals client_facturation["crm_titulaireid"]
                             join contrat in ctx.CreateQuery("crm_contrat")
                                on client_facturation["crm_titulaireid"] equals contrat["crm_account_id"]
                             where client_facturation["crm_custcode"].Equals(custCode)
                             select new CrmClient
                             {
                                 AccountId = (string) party["accountid"].ToString(),
                                 CustCode = (string) client_facturation["crm_custcode"],
                                 CustomerId = (string) client_facturation["crm_customer_id"],
                                 Contrat = (string) contrat["crm_coid"],
                                 Msisdn = (string)contrat["crm_msisdn"]
                             };

            try
            {
                client = partyQuery.First();              
            }
            catch (Exception Ex)
            {
                MiscHelper.PrintMessage($"GetCrmClientFormCustCode : {Ex.Message}");
            }           

            return client;
        }
        #endregion // GetAccountGuidFromCustCode

        #region AddCreateIncidentWorkerTask
        public void AddCreateIncidentWorkerTask(List<Task<bool>> taskList, int currentIndex, int maxCount)
        {
            var task = this.CreateIncident(currentIndex, maxCount, this.IsCallingResolve);
            taskList.Add(task);
        }
        #endregion // AddCreateIncidentWorkerTask

        #region AddResolveIncidentWorkerTask
        public void AddResolveIncidentWorkerTask(List<Task<bool>> taskList, Guid guidDemande, int currentIndex, int maxCount)
        {
            var task = this.Resolve(currentIndex, maxCount, guidDemande, true);
            taskList.Add(task);
        }
        #endregion // AddResolveIncidentWorkerTask

        #region Resolve
        public async Task<bool> Resolve(int currentIndex, int maxCount, Guid incidentId, bool isPrintingProgress)
        {
            bool isOperationSuccessfull = false;

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
                ctx.Execute(closeIncidentReq);
                if (isPrintingProgress)
                    MiscHelper.DisplayProgression(maxCount);

                isOperationSuccessfull = true;
            }
            catch (Exception ex)
            {
                var message = string.Format($"Could not close the incident : {incidentId} : {ex.Message} ");
                MiscHelper.PrintMessage(message);
            }

            return isOperationSuccessfull;
        }
        #endregion // Resolve

        #region Process
        public void Process()
        {
            foreach (var headerKey in IncidentReadDict.Keys)
            {
                switch (headerKey)
                {
                    case HEADER_STATUT:
                        {
                            var statut = IncidentReadDict[headerKey];
                            if (statut == VAUE_STATUT_RESOLVED)
                                this.IsCallingResolve = true;
                        }
                        break;
                    case HEADER_ACCOUNT:
                        {
                            var custCode = IncidentReadDict[headerKey];
                            var client = GetCrmClientFormCustCode(custCode);
                            if (client == null)
                            {
                                MiscHelper.PrintMessage($"GUID for custcode : {custCode} could not be found in DCRM");
                            }
                            else
                            {
                                IncidentWriteDict[HEADER_ACCOUNT_ID] = client.AccountId;
                                IncidentWriteDict[HEADER_CUSTCODE] = client.CustCode;
                                IncidentWriteDict[HEADER_CUSTOMER_ID] = client.CustomerId;
                                IncidentWriteDict[HEADER_CONTRAT] = client.Contrat;
                                IncidentWriteDict[HEADER_MSISDN] = client.Msisdn;
                            }
                        }
                        break;

                    case HEADER_PROPRIETAIRE:
                        {
                            var teamName = IncidentReadDict[headerKey];
                            var guid = GetGuidFromEntity("team", "teamid", "name", teamName);
                            if (string.IsNullOrEmpty(guid))
                            {
                                MiscHelper.PrintMessage($"GUID for team : {teamName} could not be found in DCRM");
                            }
                            else
                            {
                                IncidentWriteDict[HEADER_OWNER] = guid;
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

        #region CreateIncident
        public async Task<bool> CreateIncident(int currentIndex, int maxCount, bool isCallingResolve)
        {
            var isOperationSuccessfull = false;

            var ctx = DcrmConnectorFactory.GetContext();
            var dcrmConnector = DcrmConnectorFactory.Get();
            var srv = dcrmConnector.GetService();
            Guid incidentId = Guid.Empty;

            var createIncidentTimer = new MiscHelper();
            createIncidentTimer.StartTimeWatch();

            var accountId = this.IncidentWriteDict[HEADER_ACCOUNT_ID];
            if (string.IsNullOrEmpty(accountId))
            {
                var accountNumber = this.IncidentWriteDict[HEADER_CUSTOMER_ID];
                MiscHelper.PrintMessage($"GUID for account : {accountNumber} won't be created");

                return isOperationSuccessfull;
            }

            var targetIncident = new Entity("incident");
            targetIncident["customerid"] = new EntityReference(Account.EntityLogicalName, new Guid(accountId));

            targetIncident["crm_custcode"] = this.IncidentWriteDict[HEADER_CUSTCODE];
            targetIncident["crm_customer_id"] = this.IncidentWriteDict[HEADER_CUSTOMER_ID];
            targetIncident["crm_contrat"] = this.IncidentWriteDict[HEADER_CONTRAT];
            targetIncident["crm_numero_ligne_concernee"] = this.IncidentWriteDict[HEADER_MSISDN];

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

            var teamId = new Guid(this.IncidentWriteDict[HEADER_OWNER]);
            targetIncident["ownerid"] = new EntityReference(Team.EntityLogicalName, teamId);

           

            await Task<bool>.Run(() =>
            {
                incidentId = srv.Create(targetIncident);
                if (isCallingResolve)
                    isOperationSuccessfull = Resolve(currentIndex, maxCount, incidentId, false).Result;

                MiscHelper.DisplayProgression(maxCount);
            });

            createIncidentTimer.StopTimeWatch();

            return isOperationSuccessfull;
        }
        #endregion // CreateIncident

        #region GetGuid
        public Guid GetGuid()
        {
            return new Guid(IncidentReadDict[HEADER_GUID_DEMANDE]);
        }
        #endregion // GetGuid
    }
}
