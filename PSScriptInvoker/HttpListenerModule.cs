using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading;
using System.Threading.Tasks;

namespace PSScriptInvoker
{
    class HttpListenerModule
    {
        private PSScriptExecutor scriptExecutor;

        private string baseUrl;
        private string authToken;

        private HttpListener server;
        private bool isServerStopped;

        public HttpListenerModule(PSScriptExecutor scriptExecutor, string baseUrl, string authToken)
        {
            this.scriptExecutor = scriptExecutor;
            this.baseUrl = baseUrl;
            this.authToken = authToken ?? "";
            server = new HttpListener();
            server.Prefixes.Add(baseUrl);
            isServerStopped = false;
        }

        /// <summary>
        /// See also here: https://web.archive.org/web/20090720052829/http://www.switchonthecode.com/tutorials/csharp-tutorial-simple-threaded-tcp-server
        /// </summary>
        public async Task StartServerThreadAsync()
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
                        PSScriptInvoker.logWarning(msg);
                    }
                    throw ex;
                }

                PSScriptInvoker.logInfo("HttpListener successfully started. Now waiting for requests...");

                while (!isServerStopped)
                {
                    HttpListenerContext context = await server.GetContextAsync().ConfigureAwait(false);

                    Thread handleRequestThread = new Thread(new ParameterizedThreadStart(handleRequest));
                    handleRequestThread.Start(context);
                }
            }
            catch (Exception ex)
            {
                PSScriptInvoker.logError("Unexpected exception:\n" + ex.ToString());
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
                        PSScriptInvoker.logError(string.Format("Wrong auth token received: '{0}'. Do nothing and return 403 (access denied).", receivedAuthToken));
                        writeResponse(context.Response, "ERROR: Wrong auth token. Access denied!", 403);
                    }
                    else
                    {
                        // Get the URI segments
                        string[] segments = request.Url.Segments;
                        // See here for more information about URI components: https://tools.ietf.org/html/rfc3986#section-3

                        // Get parameters
                        Dictionary<String, String> parameters;
                        string body = "";
                        if (request.HasEntityBody)
                        {
                            body = getRequestBody(request);
                            parameters = JsonConvert.DeserializeObject<Dictionary<String, String>>(body);
                        }
                        else
                        {
                            string query = request.Url.Query;
                            try
                            {
                                // Parse the query string variables into a dictionary.
                                parameters = parseUriQuery(query);
                            }
                            catch (Exception ex)
                            {
                                PSScriptInvoker.logError("Unexpected exception while parsing request segments and query:\n" + ex.ToString());
                                writeResponse(context.Response, "ERROR: URL not valid: " + request.Url.ToString(), 400);
                                return;
                            }
                        }

                        // Execute the appropriate script.
                        Dictionary<String, String> scriptOutput = scriptExecutor.executePowershellScriptByHttpSegments(segments, parameters);

                        // Get output variables
                        scriptOutput.TryGetValue("exitCode", out string exitCode);
                        scriptOutput.TryGetValue("result", out string result);

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
                PSScriptInvoker.logError("Unexpected exception while processing request:\n" + ex.ToString());
                writeResponse(context.Response, ex.ToString(), 500);
            }
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
                        PSScriptInvoker.logWarning(msg);
                    }
                }
            }
            return queryDict;
        }

        private string getRequestBody(HttpListenerRequest request)
        {
            string body = "";
            try
            {
                StreamReader stream = new StreamReader(request.InputStream);
                body = stream.ReadToEnd();
            }
            catch (Exception ex)
            {
                PSScriptInvoker.logError("Exception while reading body input stream: " + ex.ToString());
            }
            return body;
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
                PSScriptInvoker.logError("Exception while writing response stream: " + ex.ToString());
                response.StatusCode = 500;
            }
            finally
            {
                response.Close();
            }
        }

        public void stopModule()
        {
            isServerStopped = true;
            server.Stop();
        }
    }
}
