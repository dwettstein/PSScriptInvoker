using Microsoft.PowerShell;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

namespace PSScriptInvoker
{
    public class PSScriptExecutor
    {
        private string pathToScripts;
        private string[] modulesToLoad;
        private string psExecutionPolicy;
        private string psOutputDelimiter;

        private RunspacePool runspacePool;
        private const int MIN_RUNSPACES = 4;

        public PSScriptExecutor(string pathToScripts, string[] modulesToLoad, string psExecutionPolicy, string psOutputDelimiter)
        {
            this.pathToScripts = pathToScripts;
            this.modulesToLoad = modulesToLoad;
            this.psExecutionPolicy = psExecutionPolicy;
            this.psOutputDelimiter = psOutputDelimiter;

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
                PSScriptInvoker.logInfo("Setting custom Powershell Execution Policy: " + psExecutionPolicy);
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
                        PSScriptInvoker.logWarning("Given custom Powershell Execution Policy is unknown: " + psExecutionPolicy + ". Only one of the following custom policies is allowed: AllSigned, Bypass, RemoteSigned, Restricted, Undefined, Unrestricted. Set to policy 'Default'.");
                        initialSession.ExecutionPolicy = ExecutionPolicy.Default;
                        break;
                }
            }

            // This loads the InitialStateSession for all instances
            // Note you can set the minimum and maximum number of runspaces as well
            // Note that without setting the minimum and maximum number of runspaces, it will use 1 as default for both:
            // https://docs.microsoft.com/en-us/dotnet/api/system.management.automation.runspaces.runspacefactory.createrunspacepool?view=powershellsdk-1.1.0
            // See here: https://stackoverflow.com/a/24358855
            runspacePool = RunspaceFactory.CreateRunspacePool(initialSession);
            runspacePool.SetMinRunspaces(MIN_RUNSPACES);
            runspacePool.SetMaxRunspaces(int.MaxValue);
            runspacePool.ThreadOptions = PSThreadOptions.UseNewThread;
            runspacePool.Open();
        }

        public void closeRunspacePool()
        {
            runspacePool.Close();
            runspacePool.Dispose();
        }

        public string getScriptPath(string[] httpSegments)
        {
            string scriptName = httpSegments[httpSegments.Length - 1].Replace("/", "");
            string scriptPath = "";
            for (int i = 0; i < httpSegments.Length - 1; i++)
            {
                scriptPath += httpSegments[i].Replace("/", "") + "\\";
            }
            return (pathToScripts + scriptPath + scriptName + ".ps1");
        }

        public string appendParameters(string scriptPath, Dictionary<String, String> parameters)
        {
            string script = scriptPath;
            foreach (string key in parameters.Keys)
            {
                string value = "";
                parameters.TryGetValue(key, out value);
                script += (" -" + key + " \"" + value + "\"");
            }
            return script;
        }

        public Dictionary<String, String> executePSScriptByHttpSegments(string[] httpSegments, Dictionary<String, String> parameters)
        {
            string script = getScriptPath(httpSegments);
            script = appendParameters(script, parameters);
            return executePSScript(script);
        }

        /**
         * See also here: http://stackoverflow.com/a/527644
         */
        public Dictionary<String, String> executePSScript(string script)
        {
            string msg = "Executing Powershell script '" + script + "'...";
            PSScriptInvoker.logInfo(msg);

            string exitCode = "";
            string result = "";
            Dictionary<String, String> output = new Dictionary<String, String>();
            Collection<PSObject> results = new Collection<PSObject>();

            try
            {
                PowerShell ps = PowerShell.Create();
                ps.AddScript(script);
                ps.RunspacePool = runspacePool;
                results = ps.Invoke();

                if (ps.HadErrors)
                {
                    exitCode = "1";
                    results.Add(new PSObject((object)ps.Streams.Error));
                }
                else
                {
                    exitCode = "0";
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
                PSScriptInvoker.logError("Exception occurred in Powershell script '" + script + "':\n" + psEx.ToString());
                results.Add(new PSObject((object)psEx.Message));
                exitCode = "1";
            }
            catch (Exception ex)
            {
                PSScriptInvoker.logError("Unexpected exception while invoking Powershell script '" + script + "':\n" + ex.ToString());
                results.Add(new PSObject((object)ex.Message));
                exitCode = "1";
            }

            try
            {
                for (int i = 0; i < results.Count; i++)
                {
                    result += (results[i] != null ? results[i].ToString() : "null");
                    if (i < results.Count - 1) // Avoid delimiter at the end.
                        result += psOutputDelimiter;
                }
            }
            catch (Exception ex)
            {
                PSScriptInvoker.logError("Failed to get result (" + results.Count + " items) of Powershell script '" + script + "':\n" + ex.ToString());
            }
            
            PSScriptInvoker.logInfo(string.Format("Executed script was: {0}. Exit code: {1}, output:\n{2}", script, exitCode, result));

            output.Add("exitCode", exitCode);
            output.Add("result", result);
            return output;
        }
    }
}
