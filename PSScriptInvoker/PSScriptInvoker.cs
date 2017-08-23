using System;
using System.Configuration;
using System.Diagnostics;
using System.ServiceProcess;
using System.Threading.Tasks;

namespace PSScriptInvoker
{
    public partial class PSScriptInvoker : ServiceBase
    {
        private const string EVENT_LOG_SOURCE = "PSScriptInvoker";
        private const string EVENT_LOG = "Application";

        // App settings
        private string httpBaseUrl;
        private string httpAuthToken;
        private string pathToScripts;
        private string[] modulesToLoad;
        private string psExecutionPolicy;

        private string rabbitMqBaseUrl;
        private string rabbitMqUsername;
        private string rabbitMqPassword;
        private string rabbitMqRequestQueueName;
        private string rabbitMqResponseQueueName;

        // Class variables
        private PSScriptExecutor scriptExecutor = null;
        private HttpListenerModule httpModule = null;
        private RabbitMqModule rabbitMqModule = null;

        public PSScriptInvoker()
        {
            InitializeComponent();

            try
            {
                if (!EventLog.SourceExists(EVENT_LOG_SOURCE))
                    EventLog.CreateEventSource(EVENT_LOG_SOURCE, EVENT_LOG);
            }
            catch (System.Security.SecurityException ex)
            {
                string msg = string.Format("Could not initialise the EventLog! Please initialise it by running the following command in Powershell as an administrator: \nNew-EventLog -Source {0} -LogName {1}\nException: {2}", EVENT_LOG_SOURCE, EVENT_LOG, ex.ToString());
                throw ex;
            }

            // Read settings from app configuration.
            httpBaseUrl = readAppSetting("baseUrl");
            httpAuthToken = readAppSetting("authToken");
            pathToScripts = readAppSetting("pathToScripts");
            string modulesToLoadString = readAppSetting("modulesToLoad");
            modulesToLoad = string.IsNullOrEmpty(modulesToLoadString) ? new string[0] : modulesToLoadString.Split(',');
            psExecutionPolicy = readAppSetting("psExecutionPolicy");
        }

        public async Task StartPSScriptInvoker()
        {
            // A shortcut to run the application for debugging within Visual Studio.
            Task startTask = OnStartTask(new string[0]);
            if (startTask != null)
                await startTask;
            else
                await new Task(new Action(waitTask));
        }

        private void waitTask()
        {
            Console.ReadLine(); // Prevent Visual Studio from exiting (just used for debugging within Visual Studio).
        }

        protected override void OnStart(string[] args)
        {
            OnStartTask(args);
            // Don't await task here in order for the service to finish starting successfully.
        }

        private Task OnStartTask(string[] args)
        {
            logInfo("Initializing service with required Powershell modules: " + String.Join(", ", modulesToLoad));
            scriptExecutor = new PSScriptExecutor(pathToScripts, modulesToLoad, psExecutionPolicy);

            if (!String.IsNullOrEmpty(httpBaseUrl))
            {
                logInfo("Initializing HTTP module on URI: " + httpBaseUrl);
                httpModule = new HttpListenerModule(scriptExecutor, httpBaseUrl, httpAuthToken);
            }
            else
            {
                string msg = string.Format("Skip HTTP initialization, some config values are missing. baseUrl: {0}", httpBaseUrl);
                logWarning(msg);
            }

            string rabbitMqBaseUrl = readAppSetting("rabbitMqBaseUrl");
            string rabbitMqUsername = readAppSetting("rabbitMqUsername");
            string rabbitMqPassword = readAppSetting("rabbitMqPassword");
            string rabbitMqRequestQueueName = readAppSetting("rabbitMqRequestQueueName");
            string rabbitMqResponseQueueName = readAppSetting("rabbitMqResponseQueueName");
            if (!String.IsNullOrEmpty(rabbitMqBaseUrl) && !String.IsNullOrEmpty(rabbitMqRequestQueueName) && !String.IsNullOrEmpty(rabbitMqResponseQueueName))
            {
                logInfo(string.Format("Initializing RabbitMQ module on URI: {0} with requestQueue: {1} and responseQueue: {2}", rabbitMqBaseUrl, rabbitMqRequestQueueName, rabbitMqResponseQueueName));
                rabbitMqModule = new RabbitMqModule(scriptExecutor, rabbitMqBaseUrl, rabbitMqUsername, rabbitMqPassword, rabbitMqRequestQueueName, rabbitMqResponseQueueName);
            }
            else
            {
                string msg = string.Format("Skip RabbitMQ initialization, some config values are missing. baseUrl: {0}, requestQueueName: {1}, responseQueueName: {2}", rabbitMqBaseUrl, rabbitMqRequestQueueName, rabbitMqResponseQueueName);
                logWarning(msg);
            }

            if (httpModule == null && (rabbitMqModule == null || !rabbitMqModule.isConnected()))
            {
                string msg = "No module was initialized. Service is not usable...";
                logError(msg);
                throw new Exception(msg);
            }
            else
            {
                logInfo("The service has been initialized successfully.");
            }

            return (httpModule != null ? httpModule.StartServerThreadAsync() : null);
        }

        protected override void OnStop()
        {
            if (httpModule != null)
                httpModule.stopModule();
            if (rabbitMqModule != null)
                rabbitMqModule.stopModule();
            if (scriptExecutor != null)
                scriptExecutor.closeRunspacePool();
            logInfo("The service has been stopped.");
        }

        private string readAppSetting(string key)
        {
            string result = "";
            try
            {
                var appSettings = ConfigurationManager.AppSettings;
                result = appSettings[key] ?? null;
            }
            catch (Exception ex)
            {
                logError("Exception while reading app setting: " + ex.ToString());
            }
            return result;
        }

        private static void logMsg(string message, EventLogEntryType level)
        {
            EventLog.WriteEntry(EVENT_LOG_SOURCE, message, level);
        }

        public static void logInfo(string message)
        {
            logMsg(message, EventLogEntryType.Information);
        }

        public static void logWarning(string message)
        {
            logMsg(message, EventLogEntryType.Warning);
        }

        public static void logError(string message)
        {
            logMsg(message, EventLogEntryType.Error);
        }

        public PSScriptExecutor getScriptExecutor()
        {
            return scriptExecutor;
        }


    }
}
