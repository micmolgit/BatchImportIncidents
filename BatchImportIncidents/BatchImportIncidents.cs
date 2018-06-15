using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using DCRM_Utils;
using Microsoft.Extensions.Configuration;
using System.Threading;

namespace BatchImportIncidents
{
    class BatchImportIncidents
    {
        #region Properties
        public readonly bool _IsCallingResolve = false;
        public readonly bool _IsCallingImport = false;
        public bool IsDebugMode { get; set; }
        public string OutputDir { get; set; }
        public string OutputFile { get; set; }
        public string OutputFilePath { get; set; }
        public string InputFilePath { get; set; }
        public string Command { get; set; }
        public string AccountHeader { get; set; }
        public string ClientHeader { get; set; }

        #endregion // Properties

        #region Constructors
        public BatchImportIncidents(string inputFilePath, string command, bool isCallingResolve, bool isCallingImport)
        {
            InputFilePath = inputFilePath;
            Command = command;
            _IsCallingResolve = isCallingResolve;
            _IsCallingImport = isCallingImport;
        }

        public BatchImportIncidents()
        {
            this.LoadConfiguration();
        }

        #endregion // Constructor

        #region Terminate
        public void Terminate()
        {
            MiscHelper.PauseExecution();
        }
        #endregion //Terminate

        #region GetHeader
        private string GetHeaderLine(IList<string> headerItems)
        {
            var sbHeader = new StringBuilder();
            for (var i = 0; i < headerItems.Count; i++)
            {
                var curHeader = headerItems[i];
                var field = string.Format($"{headerItems[i]};");

                if (curHeader == AccountHeader)
                {
                    sbHeader.Append(ClientHeader);
                }
                else
                {
                    sbHeader.Append(field);
                }
            }
            var outputLine = sbHeader.ToString();

            return (outputLine);
        }
        #endregion // GetHeader

        #region GetDemandeFromLine
        public Demande GetDemandeFromLine(IList<string> header, IList<string> line)
        {
            var demande = new Demande();

            for (var i = 0; i < header.Count; i++)
            {
                if (!string.IsNullOrEmpty(line[i]))
                {
                    var field = string.Format($"{line[i]};");
                    var curHeader = header[i];
                    var curValue = line[i];

                    // We're saving the value of the field into a dictionary for later purpose
                    demande.SetFieldValue(curHeader, curValue);
                }
            }
            return demande;
        }
        #endregion // GetDemandeFromLine

        #region Process
        public async Task<int> Process()
        {
            var entriesCpt = 0;

            if (_IsCallingImport)
                entriesCpt = await ProcessImportAsync();
            else if (_IsCallingResolve)
                entriesCpt = await ProcessResolveAsync();

            return entriesCpt;
        }
        #endregion // Process

        #region ProcessResolveAsync
        public async Task<int> ProcessResolveAsync()
        {
            var entriesCpt = 0;

            using (var dataReadStream = File.OpenRead(InputFilePath))
            using (var dataReader = new StreamReader(dataReadStream))

            using (var countReadStream = File.OpenRead(InputFilePath))
            using (var countReader = new StreamReader(countReadStream))
            {
                // parsing .csv input file
                var dataRead = CsvParser.ParseHeadAndTail(dataReader, ';', '"');
                var countRead = CsvParser.ParseHeadAndTail(countReader, ';', '"');

                var countQuery = countRead.Item2;
                var dataQuery = dataRead.Item2;

                MiscHelper.PrintMessage("Counting the number of incident to resolve...");
                var incidentCountTimer = new MiscHelper();
                incidentCountTimer.StartTimeWatch();
                var maxCount = countQuery.Count();
                incidentCountTimer.StopTimeWatch();
                MiscHelper.PrintMessage($"\n\n{maxCount} were scanned in {incidentCountTimer.GetDuration()}");

                MiscHelper.PrintMessage("Connecting do DCRM...");
                DcrmConnectorFactory.GetContext();

                var dataHeader = dataRead.Item1;

                // writting header to .csv output file
                var headerLine = GetHeaderLine(dataHeader);

                var resolvencidentThreads = new List<Thread>();

                var ResolveTimer = new MiscHelper();
                ResolveTimer.StartTimeWatch();

                var resolveIncidentTasks = new List<Task<bool>>();

                // read every lines of the .csv source file and translate them into the .csv output file
                foreach (var line in dataQuery)
                {
                    if (line.Count > 0)
                    {
                        var demande = GetDemandeFromLine(dataHeader, line);
                        var guidDemande = demande.GetGuid();
                        demande.AddResolveIncidentWorkerTask(resolveIncidentTasks, guidDemande, ++entriesCpt, maxCount);
                    }
                }

                // We need to wait for all tasks to be terminated
                MiscHelper.PrintMessage($"\nProcessing the {maxCount} incidents, please wait ...");

                await System.Threading.Tasks.Task.WhenAll(resolveIncidentTasks);

                MiscHelper.PrintMessage("\nAll The operation completed successfully");

                ResolveTimer.StopTimeWatch();
                MiscHelper.PrintMessage($"\n\n{entriesCpt} incidents were resolved in {ResolveTimer.GetDuration()}");
            }

            return entriesCpt;
        }
        #endregion // ProcessResolveAsync

        #region ProcessImportAsync
        public async Task<int> ProcessImportAsync()
        {
            var entriesCpt = 0;

            using (var dataReadStream = File.OpenRead(InputFilePath))
            using (var dataReader = new StreamReader(dataReadStream))

            using (var countReadStream = File.OpenRead(InputFilePath))
            using (var countReader = new StreamReader(countReadStream))
            {
                // parsing .csv input file
                var dataRead = CsvParser.ParseHeadAndTail(dataReader, ';', '"');
                var countRead = CsvParser.ParseHeadAndTail(countReader, ';', '"');

                var dataQuery = dataRead.Item2;
                var countQuery = countRead.Item2;

                MiscHelper.PrintMessage("Counting the number of incidents to import...");
                var incidentCountTimer = new MiscHelper();
                incidentCountTimer.StartTimeWatch();
                var maxCount = countQuery.Count();
                incidentCountTimer.StopTimeWatch();
                MiscHelper.PrintMessage($"{maxCount} were scanned in {incidentCountTimer.GetDuration()}");

                MiscHelper.PrintMessage("Connecting do DCRM...");
                DcrmConnectorFactory.GetContext();

                var dataHeader = dataRead.Item1;

                var headerLine = GetHeaderLine(dataHeader);

                var createIncidentTasks = new List<Task<bool>>();

                var ImportTimer = new MiscHelper();
                ImportTimer.StartTimeWatch();

                MiscHelper.PrintMessage($"\nProcessing the {maxCount} incidents, please wait ...");

                // read every lines of the .csv source file and translate them into the .csv output file
                foreach (var line in dataQuery)
                {
                    if (line.Count > 0)
                    {
                        var demande = GetDemandeFromLine(dataHeader, line);
                        demande.Process();
                        demande.AddCreateIncidentWorkerTask(createIncidentTasks, ++entriesCpt, maxCount);
                    }
                }                

                // We need to wait for all tasks to be terminated
                await System.Threading.Tasks.Task.WhenAll(createIncidentTasks);

                MiscHelper.PrintMessage("\nAll The operation completed successfully");

                ImportTimer.StopTimeWatch();
                MiscHelper.PrintMessage($"\n\n{entriesCpt} incidents were created in {ImportTimer.GetDuration()}");
            }

            return entriesCpt;
        }
        #endregion // ProcessImportAsync

        #region GetConfiguration
        private IConfigurationRoot GetConfiguration()
        {
            var builder = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json");
            var configuration = builder.Build();

            return configuration;
        }
        #endregion // GetConfiguration

        #region LoadConfiguration
        private void LoadConfiguration()
        {
            var configuration = GetConfiguration();

            ClientHeader = configuration["ClientHeader"];
            AccountHeader = configuration["AccountHeader"];
            OutputDir = configuration["OutputDir"];
            OutputFile = configuration["OutputFile"];
            OutputFilePath = string.Format($@"{OutputDir}\{OutputFile}");
            IsDebugMode = configuration["IsDebugMode"] == "true";
        }
        #endregion // LoadConfiguration
    }
}
