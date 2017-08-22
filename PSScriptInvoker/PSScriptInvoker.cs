using System;
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

            EventLog.WriteEntry(EVENT_LOG_SOURCE, "The service has been started.", EventLogEntryType.Information);
            return StartServerThreadAsync();
        }

        protected override void OnStop()
        {
            isServerStopped = true;
            server.Stop();
            server = null;
            //if (serverThread.IsAlive)
            //    serverThread.Abort();
            runspacePool.Close();
            EventLog.WriteEntry(EVENT_LOG_SOURCE, "The service has been stopped.", EventLogEntryType.Information);
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
                        Dictionary<String, String> scriptOutput = executePowershellScript(scriptPath, scriptName, queryDict);

                        // Get output variables
                        string returnCode = "";
                        string result = "";
                        scriptOutput.TryGetValue("returnCode", out returnCode);
                        scriptOutput.TryGetValue("result", out result);

                        string msg = string.Format("Executed script was: {0}. Return code: {1}, output:\n{2}", request.Url.ToString(), returnCode, result);
                        Console.WriteLine(msg);
                        EventLog.WriteEntry(EVENT_LOG_SOURCE, msg, EventLogEntryType.Information);

                        if (returnCode == "0")
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

        /**
         * See here http://stackoverflow.com/a/527644
         */
        private Dictionary<String, String> executePowershellScript(string scriptPath, string scriptName, Dictionary<String, String> inputs)
        {
            string fullScriptPath = pathToScripts + scriptPath + scriptName + ".ps1";
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
            IList errors = new ArrayList();

            try
            {
                PowerShell ps = PowerShell.Create();
                ps.AddScript(fullScriptPath);
                ps.RunspacePool = runspacePool;
                results = ps.Invoke();

                if (results.Count > 0)
                {
                    output.Add("returnCode", "0");
                }
                else
                {
                    output.Add("returnCode", "1");
                    if (errors.Count > 0)
                    {
                        results.Add(new PSObject((object)errors[0]));
                    }
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
                output.Add("returnCode", "1");
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry(EVENT_LOG_SOURCE, "Unexpected exception while invoking Powershell script '" + fullScriptPath + "':\n" + ex.ToString(), EventLogEntryType.Error);
                results.Add(new PSObject((object)ex.Message));
                output.Add("returnCode", "1");
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

        private string readAppSetting(string key)
        {
            string result = "";
            try
            {
                var appSettings = ConfigurationManager.AppSettings;
                result = appSettings[key] ?? "Not Found";
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry(EVENT_LOG_SOURCE, "Exception while reading app setting: " + ex.ToString(), EventLogEntryType.Error);
            }
            return result;
        }
    }
}
