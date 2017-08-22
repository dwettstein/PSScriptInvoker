﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Net;
using System.ServiceProcess;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.PowerShell;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;

namespace PSScriptInvoker
{
    public partial class PSScriptInvoker : ServiceBase
    {
        private const string EVENT_LOG_SOURCE = "PSScriptInvoker";
        private const string EVENT_LOG = "Application";

        private string baseUrl;
        private string authToken;
        private string pathToScripts;
        private string[] modulesToLoad;
        private string psExecutionPolicy;

        private string rabbitMqBaseUrl;
        private string rabbitMqUsername;
        private string rabbitMqPassword;
        private string rabbitMqRequestQueueName;
        private string rabbitMqResponseQueueName;

        private ConnectionFactory rabbitmqFactory;
        private IConnection rabbitmqConnection;
        private IModel rabbitmqChannel;
        private AsyncEventingBasicConsumer rabbitmqConsumer;

        private HttpListener server;
        private bool isServerStopped;
        private RunspacePool runspacePool;
        private int MIN_RUNSPACES = 4;

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
                Console.WriteLine(msg);
                throw ex;
            }

            // Read settings from app configuration.
            baseUrl = readAppSetting("baseUrl");
            authToken = readAppSetting("authToken");
            pathToScripts = readAppSetting("pathToScripts");
            string modulesToLoadString = readAppSetting("modulesToLoad");
            modulesToLoad = string.IsNullOrEmpty(modulesToLoadString) ? new string[0] : modulesToLoadString.Split(',');
            psExecutionPolicy = readAppSetting("psExecutionPolicy");

            rabbitMqBaseUrl = readAppSetting("rabbitMqBaseUrl");
            rabbitMqUsername = readAppSetting("rabbitMqUsername");
            rabbitMqPassword = readAppSetting("rabbitMqPassword");
            rabbitMqRequestQueueName = readAppSetting("rabbitMqRequestQueueName");
            rabbitMqResponseQueueName = readAppSetting("rabbitMqResponseQueueName");
        }

        public async Task StartPSScriptInvoker()
        {
            // A shortcut to run the application for debugging within Visual Studio.
            await OnStartTask(new string[0]);
        }

        protected override void OnStart(string[] args)
        {
            OnStartTask(args);
            // Don't await task here in order for the service to finish starting successfully.
        }

        private Task OnStartTask(string[] args)
        {
            EventLog.WriteEntry(EVENT_LOG_SOURCE, "Initialising service on URI: " + baseUrl + " with required Powershell modules: " + String.Join(", ", modulesToLoad), EventLogEntryType.Information);

            // Initialise Powershell Runspace and preload necessary modules.
            // See here: http://stackoverflow.com/a/17071164
            // See here: http://nivot.org/blog/post/2010/05/03/PowerShell20DeveloperEssentials1InitializingARunspaceWithAModule
            InitialSessionState initialSession = InitialSessionState.CreateDefault2();
            if (modulesToLoad != null && modulesToLoad.Length > 0)
            {
                initialSession.ImportPSModule(modulesToLoad);
            }

            if (psExecutionPolicy != "None")
            {
                EventLog.WriteEntry(EVENT_LOG_SOURCE, "Setting custom Powershell Execution Policy: " + psExecutionPolicy, EventLogEntryType.Information);
                switch (psExecutionPolicy)
                {
                    case "AllSigned":
                        initialSession.ExecutionPolicy = ExecutionPolicy.AllSigned;
                        break;
                    case "Bypass":
                        initialSession.ExecutionPolicy = ExecutionPolicy.Bypass;
                        break;
                    case "RemoteSigned":
                        initialSession.ExecutionPolicy = ExecutionPolicy.RemoteSigned;
                        break;
                    case "Restricted":
                        initialSession.ExecutionPolicy = ExecutionPolicy.Restricted;
                        break;
                    case "Undefined":
                        initialSession.ExecutionPolicy = ExecutionPolicy.Undefined;
                        break;
                    case "Unrestricted":
                        initialSession.ExecutionPolicy = ExecutionPolicy.Unrestricted;
                        break;
                    default:
                        EventLog.WriteEntry(EVENT_LOG_SOURCE, "Given custom Powershell Execution Policy is unknown: " + psExecutionPolicy + ". Only one of the following custom policies is allowed: AllSigned, Bypass, RemoteSigned, Restricted, Undefined, Unrestricted. Set to policy 'Default'.", EventLogEntryType.Warning);
                        initialSession.ExecutionPolicy = ExecutionPolicy.Default;
                        break;
                }
            }

            // This loads the InitialStateSession for all instances
            // Note you can set the minimum and maximum number of runspaces as well
            // See here: https://stackoverflow.com/a/24358855
            runspacePool = RunspaceFactory.CreateRunspacePool(initialSession);
            runspacePool.SetMinRunspaces(MIN_RUNSPACES);
            runspacePool.Open();

            server = new HttpListener();
            server.Prefixes.Add(baseUrl);
            isServerStopped = false;

            initializeRabbitMQ(rabbitMqBaseUrl, rabbitMqRequestQueueName, rabbitMqUsername, rabbitMqPassword);

            EventLog.WriteEntry(EVENT_LOG_SOURCE, "The service has been started.", EventLogEntryType.Information);
            return StartServerThreadAsync();
        }

        protected override void OnStop()
        {
            isServerStopped = true;
            server.Stop();
            server = null;
            runspacePool.Close();
            EventLog.WriteEntry(EVENT_LOG_SOURCE, "The service has been stopped.", EventLogEntryType.Information);
        }

        private void initializeRabbitMQ(string baseUrl, string queueName, string username, string password)
        {
            try
            {
                if (!String.IsNullOrEmpty(baseUrl) && !String.IsNullOrEmpty(queueName))
                {
                    rabbitmqFactory = new ConnectionFactory() { Uri = new Uri(baseUrl), UserName = username, Password = password, DispatchConsumersAsync = true };
                    rabbitmqFactory.AutomaticRecoveryEnabled = true;
                    rabbitmqConnection = rabbitmqFactory.CreateConnection();
                    rabbitmqChannel = rabbitmqConnection.CreateModel();
                    rabbitmqConsumer = new AsyncEventingBasicConsumer(rabbitmqChannel);
                    rabbitmqConsumer.Received += async (sender, args) =>
                    {
                        await Task.Yield(); // Force async execution.

                        var body = args.Body;
                        var message = System.Text.Encoding.UTF8.GetString(body);

                        EventLog.WriteEntry(EVENT_LOG_SOURCE, string.Format("Received RabbitMQ message (deliveryTag: {0}):\n{1}", args.DeliveryTag, message), EventLogEntryType.Information);

                        handleMessage(message);

                        lock (rabbitmqChannel)
                        {
                            rabbitmqChannel.BasicAck(deliveryTag: args.DeliveryTag, multiple: false);
                        }
                    };

                    rabbitmqChannel.BasicConsume(queue: queueName,
                                 autoAck: false,
                                 consumer: rabbitmqConsumer);

                    EventLog.WriteEntry(EVENT_LOG_SOURCE, "Message consumer successfully started. Now waiting for requests in queue " + queueName + "...", EventLogEntryType.Information);
                }
                else
                {
                    string msg = string.Format("Skip RabbitMQ initialization, some config values are missing. baseUrl: {0}, queueName: {1}", baseUrl, queueName);
                    Console.WriteLine(msg);
                    EventLog.WriteEntry(EVENT_LOG_SOURCE, msg, EventLogEntryType.Warning);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                EventLog.WriteEntry(EVENT_LOG_SOURCE, "Unexpected exception while setting up RabbitMQ connection:\n" + ex.ToString(), EventLogEntryType.Error);
            }
        }

        /// <summary>
        /// See also here: https://web.archive.org/web/20090720052829/http://www.switchonthecode.com/tutorials/csharp-tutorial-simple-threaded-tcp-server
        /// </summary>
        private async Task StartServerThreadAsync()
        {
            try
            {
                // Start listening for client requests.
                try
                {
                    server.Start();
                }
                catch (HttpListenerException ex)
                {
                    if (ex.ErrorCode == 5) // Access is denied
                    {
                        string msg = string.Format("You need to reserve the URI by running the following command as an administrator: \nnetsh http add urlacl url={0} user={1}\\{2} listen=yes", baseUrl, Environment.GetEnvironmentVariable("USERDOMAIN"), Environment.GetEnvironmentVariable("USERNAME"));
                        Console.WriteLine(msg);
                        EventLog.WriteEntry(EVENT_LOG_SOURCE, msg, EventLogEntryType.Warning);
                    }
                    throw ex;
                }

                EventLog.WriteEntry(EVENT_LOG_SOURCE, "HttpListener successfully started. Now waiting for requests...", EventLogEntryType.Information);

                while (!isServerStopped)
                {
                    HttpListenerContext context = await server.GetContextAsync().ConfigureAwait(false);

                    Thread handleRequestThread = new Thread(new ParameterizedThreadStart(handleRequest));
                    handleRequestThread.Start(context);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                EventLog.WriteEntry(EVENT_LOG_SOURCE, "Unexpected exception:\n" + ex.ToString(), EventLogEntryType.Error);
            }
        }

        private void handleRequest(object input)
        {
            HttpListenerContext context = (HttpListenerContext)input;

            // A request was arrived. Get the object.
            HttpListenerRequest request = context.Request;

            try
            {
                string receivedAuthToken = request.Headers.Get("Authorization");
                if (!string.IsNullOrEmpty(authToken) && string.IsNullOrEmpty(receivedAuthToken))
                {
                    writeResponse(context.Response, "ERROR: Authorization header missing!", 401);
                }
                else
                {
                    if (!string.IsNullOrEmpty(authToken) && !receivedAuthToken.Equals(authToken))
                    {
                        EventLog.WriteEntry(EVENT_LOG_SOURCE, string.Format("Wrong auth token received: '{0}'. Do nothing and return 403 (access denied).", receivedAuthToken), EventLogEntryType.FailureAudit);
                        writeResponse(context.Response, "ERROR: Wrong auth token. Access denied!", 403);
                    }
                    else
                    {
                        string query;
                        string[] segments;
                        Dictionary<String, String> queryDict;
                        try
                        {
                            // See here for more information about URI components: https://tools.ietf.org/html/rfc3986#section-3
                            query = request.Url.Query;
                            // Parse the query string variables into a dictionary.
                            queryDict = parseUriQuery(query);
                            //NameValueCollection queryCollection = System.Web.HttpUtility.ParseQueryString(querystring);

                            // Get the URI segments
                            segments = request.Url.Segments;
                        }
                        catch (Exception ex)
                        {
                            EventLog.WriteEntry(EVENT_LOG_SOURCE, "Unexpected exception while parsing request segments and query:\n" + ex.ToString(), EventLogEntryType.Error);
                            writeResponse(context.Response, "ERROR: URL not valid: " + request.Url.ToString(), 400);
                            return;
                        }

                        // Execute the appropriate script.
                        string scriptName = segments[segments.Length - 1].Replace("/", "");
                        string scriptPath = "";
                        for (int i = 0; i < segments.Length - 1; i++)
                        {
                            scriptPath += segments[i].Replace("/", "") + "\\";
                        }
                        string fullScriptPath = pathToScripts + scriptPath + scriptName + ".ps1";
                        Dictionary<String, String> scriptOutput = executePowershellScript(fullScriptPath, queryDict);

                        // Get output variables
                        string exitCode = "";
                        string result = "";
                        scriptOutput.TryGetValue("exitCode", out exitCode);
                        scriptOutput.TryGetValue("result", out result);

                        string msg = string.Format("Executed script was: {0}. Exit code: {1}, output:\n{2}", fullScriptPath, exitCode, result);
                        Console.WriteLine(msg);
                        EventLog.WriteEntry(EVENT_LOG_SOURCE, msg, EventLogEntryType.Information);

                        if (exitCode == "0")
                        {
                            if (string.IsNullOrEmpty(result))
                            {
                                writeResponse(context.Response, result, 204);
                            }
                            else
                            {
                                writeResponse(context.Response, result, 200);
                            }
                        }
                        else
                        {
                            writeResponse(context.Response, result, 500);
                        }
                    }
              }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry(EVENT_LOG_SOURCE, "Unexpected exception while processing request:\n" + ex.ToString(), EventLogEntryType.Error);
                writeResponse(context.Response, ex.ToString(), 500);
            }
        }

        private void handleMessage(string message)
        {
            try
            {
                // Get parameters
                Dictionary<String, String> parameters = new Dictionary<String, String>();

                // Execute the appropriate script.
                string scriptName = "";
                string scriptPath = "";
                string fullScriptPath = pathToScripts + scriptPath + scriptName + ".ps1";
                Dictionary<String, String> scriptOutput = executePowershellScript(fullScriptPath, parameters);

                // Get output variables
                string exitCode = "";
                string result = "";
                scriptOutput.TryGetValue("exitCode", out exitCode);
                scriptOutput.TryGetValue("result", out result);

                string msg = string.Format("Executed script was: {0}. Exit code: {1}, output:\n{2}", fullScriptPath, exitCode, result);
                Console.WriteLine(msg);
                EventLog.WriteEntry(EVENT_LOG_SOURCE, msg, EventLogEntryType.Information);

                if (exitCode == "0")
                {
                    if (string.IsNullOrEmpty(result))
                    {
                        writeMessage("");
                        //writeResponse(context.Response, result, 204);
                    }
                    else
                    {
                        writeMessage("");
                        //writeResponse(context.Response, result, 200);
                    }
                }
                else
                {
                    writeMessage("");
                    //writeResponse(context.Response, result, 500);
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry(EVENT_LOG_SOURCE, "Unexpected exception while processing message:\n" + ex.ToString(), EventLogEntryType.Error);
                writeMessage(ex.ToString());
                //writeResponse(context.Response, ex.ToString(), 500);
            }
        }

        /**
         * See here http://stackoverflow.com/a/527644
         */
        private Dictionary<String, String> executePowershellScript(string fullScriptPath, Dictionary<String, String> inputs)
        {
            foreach (string key in inputs.Keys)
            {
                string value = "";
                inputs.TryGetValue(key, out value);
                fullScriptPath += (" -" + key + " " + value);
            }

            string msg = "Executing Powershell script '" + fullScriptPath + "'...";
            Console.WriteLine(msg);
            EventLog.WriteEntry(EVENT_LOG_SOURCE, msg, EventLogEntryType.Information);

            Dictionary<String, String> output = new Dictionary<String, String>();
            Collection<PSObject> results = new Collection<PSObject>();

            try
            {
                PowerShell ps = PowerShell.Create();
                ps.AddScript(fullScriptPath);
                ps.RunspacePool = runspacePool;
                results = ps.Invoke();

                if (ps.HadErrors)
                {
                    output.Add("exitCode", "1");
                    results.Add(new PSObject((object)ps.Streams.Error));
                }
                else
                {
                    output.Add("exitCode", "0");
                }
            }
            catch (ActionPreferenceStopException ex)
            {
                Exception psEx = null;
                if (ex.ErrorRecord != null && ex.ErrorRecord.Exception != null)
                {
                    psEx = ex.ErrorRecord.Exception;
                }
                else
                {
                    psEx = ex;
                }
                EventLog.WriteEntry(EVENT_LOG_SOURCE, "Exception occurred in Powershell script '" + fullScriptPath + "':\n" + psEx.ToString(), EventLogEntryType.Error);
                results.Add(new PSObject((object)psEx.Message));
                output.Add("exitCode", "1");
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry(EVENT_LOG_SOURCE, "Unexpected exception while invoking Powershell script '" + fullScriptPath + "':\n" + ex.ToString(), EventLogEntryType.Error);
                results.Add(new PSObject((object)ex.Message));
                output.Add("exitCode", "1");
            }

            if (results.Count > 0)
            {
                output.Add("result", results[0].ToString());
            }
            else
            {
                output.Add("result", "");
            }
            return output;
        }

        private Dictionary<String, String> parseUriQuery(string query)
        {
            Dictionary<String, String> queryDict = new Dictionary<String, String>();
            if (!string.IsNullOrEmpty(query))
            {
                string[] queryParams = query.Split('&');
                foreach (string param in queryParams)
                {
                    string[] parts = param.Split('=');
                    if (parts.Length == 1)
                    {
                        string key = Uri.UnescapeDataString(parts[0].Trim(new char[] { '?', ' ' }));
                        queryDict.Add(key, null);
                    }
                    else if (parts.Length == 2)
                    {
                        string key = Uri.UnescapeDataString(parts[0].Trim(new char[] { '?', ' ' }));
                        string value = Uri.UnescapeDataString(parts[1].Trim());
                        queryDict.Add(key, value);
                    }
                    else
                    {
                        string msg = string.Format("Unable to parse param: {0}", param);
                        Console.WriteLine(msg);
                        EventLog.WriteEntry(EVENT_LOG_SOURCE, msg, EventLogEntryType.Warning);
                    }
                }
            }
            return queryDict;
        }

        private void writeResponse(HttpListenerResponse response, string responseText, int statusCode)
        {
            try
            {
                response.StatusCode = statusCode;
                if (!string.IsNullOrEmpty(responseText))
                {
                    Stream stream = response.OutputStream;
                    byte[] msg = System.Text.Encoding.ASCII.GetBytes(responseText);
                    stream.Write(msg, 0, msg.Length);
                    stream.Close();
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry(EVENT_LOG_SOURCE, "Exception while writing response stream: " + ex.ToString(), EventLogEntryType.Error);
                response.StatusCode = 500;
            }
            finally
            {
                response.Close();
            }
        }

        private void writeMessage(string messageText)
        {

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
                EventLog.WriteEntry(EVENT_LOG_SOURCE, "Exception while reading app setting: " + ex.ToString(), EventLogEntryType.Error);
            }
            return result;
        }
    }
}
