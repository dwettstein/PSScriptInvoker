﻿using Microsoft.PowerShell;
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

        private RunspacePool runspacePool;
        private int MIN_RUNSPACES = 4;

        public PSScriptExecutor(string pathToScripts, string[] modulesToLoad, string psExecutionPolicy)
        {
            this.pathToScripts = pathToScripts;
            this.modulesToLoad = modulesToLoad;
            this.psExecutionPolicy = psExecutionPolicy;

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
            // See here: https://stackoverflow.com/a/24358855
            runspacePool = RunspaceFactory.CreateRunspacePool(initialSession);
            runspacePool.SetMinRunspaces(MIN_RUNSPACES);
            runspacePool.Open();
        }

        /**
         * See here http://stackoverflow.com/a/527644
         */
        public Dictionary<String, String> executePowershellScript(string fullScriptPath, Dictionary<String, String> inputs)
        {
            foreach (string key in inputs.Keys)
            {
                string value = "";
                inputs.TryGetValue(key, out value);
                fullScriptPath += (" -" + key + " " + value);
            }

            string msg = "Executing Powershell script '" + fullScriptPath + "'...";
            PSScriptInvoker.logInfo(msg);

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
                PSScriptInvoker.logError("Exception occurred in Powershell script '" + fullScriptPath + "':\n" + psEx.ToString());
                results.Add(new PSObject((object)psEx.Message));
                output.Add("exitCode", "1");
            }
            catch (Exception ex)
            {
                PSScriptInvoker.logError("Unexpected exception while invoking Powershell script '" + fullScriptPath + "':\n" + ex.ToString());
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

        public void closeRunspacePool()
        {
            runspacePool.Close();
        }

        public string getPathToScripts()
        {
            return pathToScripts;
        }
    }
}