using System;
using System.ServiceModel;
using Nito.AsyncEx;
using System.Reflection;
using System.Threading.Tasks;
using DCRM_Utils;

namespace BatchImportIncidents
{
    class Program
    {
        #region Constants
        private const string STR_IMPORT = "i";
        private const string STR_RESOLVE = "r";
        #endregion // Constants

        #region PrintUsage
        public static void PrintUsage()
        {
            var exeName = Assembly.GetExecutingAssembly().GetName().Name;

            MiscHelper.PrintMessage(
              $"**************************************************************************************************************\n" +
              $"* Usage :                                                                                                    *\n" +
              $"*                                                                                                            *\n" +
              $"* - {exeName}.exe -{STR_IMPORT} <.csv path> to import incidents                                              *\n" +
              $"*                                                                                                            *\n" +
              $"* - {exeName}.exe -{STR_RESOLVE} <.csv path> to resolve incidents                                             *\n" +
              $"*                                                                                                            *\n" +
              $"**************************************************************************************************************");
            MiscHelper.PauseExecution();
        }
        #endregion // PrintUsage

        #region MainAsync
        // We are creating an async main method based on Nito.AsyncEx
        static async Task<bool> MainAsync(string[] args)
        {
            bool isOperationSuccessfull = false;

            if (args.Length < 2)
            {
                PrintUsage();
                return isOperationSuccessfull;
            }

            var command = args[0];
            var sourceFilePath = args[1];
            var isCallingResolve = false;
            var isCallingImport = false;

            BatchImportIncidents batch = null;

            try
            {
                if (command.Contains(STR_IMPORT))
                    isCallingImport = true;
                if (command.Contains(STR_RESOLVE))
                    isCallingResolve = true;

                if (!isCallingImport && !isCallingResolve)
                {
                    var message = string.Format($"The specified ({command}) command is not supported.");
                    throw new ArgumentException(message);
                }

                batch = new BatchImportIncidents(sourceFilePath, command, isCallingResolve, isCallingImport);
                var incidentCountBatch = await batch.Process();

                isOperationSuccessfull = true;
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                MiscHelper.PrintMessage("The application terminated with an error.");
                MiscHelper.PrintMessage($"Timestamp: { ex.Detail.Timestamp}");
                MiscHelper.PrintMessage($"Code: {ex.Detail.ErrorCode}");
                MiscHelper.PrintMessage($"Message: {ex.Detail.Message}");
                MiscHelper.PrintMessage($"Plugin Trace: {ex.Detail.TraceText}");
                if (ex.InnerException != null)
                    MiscHelper.PrintMessage($"Inner Fault: { ex.InnerException.Message ?? "No Inner Fault"}");
            }
            catch (System.TimeoutException ex)
            {
                MiscHelper.PrintMessage("The application terminated with an error.");
                MiscHelper.PrintMessage($"Message: {ex.Message}");
                MiscHelper.PrintMessage($"Stack Trace: {ex.StackTrace}");
                if (ex.InnerException != null)
                    MiscHelper.PrintMessage($"Inner Fault: { ex.InnerException.Message ?? "No Inner Fault"}");
            }
            catch (System.Exception ex)
            {
                MiscHelper.PrintMessage($"The application terminated with an error : {ex.Message}");

                // Display the details of the inner exception.
                if (ex.InnerException != null)
                {
                    MiscHelper.PrintMessage(ex.InnerException.Message);

                    if (ex.InnerException is FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> fe)
                    {
                        MiscHelper.PrintMessage($"Timestamp: {fe.Detail.Timestamp}");
                        MiscHelper.PrintMessage($"Code: {fe.Detail.ErrorCode}");
                        MiscHelper.PrintMessage($"Message: {fe.Detail.Message}");
                        MiscHelper.PrintMessage($"Plugin Trace: {fe.Detail.TraceText}");
                        var message = string.Format("Inner Fault: {0}",
                            null == fe.Detail.InnerFault ? "No Inner Fault" : "Has Inner Fault");
                        MiscHelper.PrintMessage(message);
                    }
                }
            }
            finally
            {
                if (batch != null)
                    batch.Terminate();
            }

            return isOperationSuccessfull;
        }
        #endregion //MainAsync

        #region Main
        static void Main(string[] args)
        {
            AsyncContext.Run(() => MainAsync(args));
            //AsyncContext.Run(() => WorkWithTasksAsync());
        }
        #endregion // Main
    }
}
