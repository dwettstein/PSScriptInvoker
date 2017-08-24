using Newtonsoft.Json;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace PSScriptInvoker
{
    class RabbitMqModule
    {
        private PSScriptExecutor scriptExecutor;

        private string baseUrl;
        private string username;
        private string password;
        private string requestQueue;
        private string responseExchange;
        private string responseRoutingKey;

        private ConnectionFactory rabbitMqFactory;
        private IConnection rabbitMqConnection;
        private IModel rabbitMqChannel;
        private AsyncEventingBasicConsumer rabbitMqConsumer;

        /**
         * See also here: https://www.rabbitmq.com/dotnet-api-guide.html
         */
        public RabbitMqModule(PSScriptExecutor scriptExecutor, string baseUrl, string username, string password, string requestQueue, string responseExchange, string responseRoutingKey)
        {
            this.scriptExecutor = scriptExecutor;
            this.baseUrl = baseUrl;
            this.username = username ?? "";
            this.password = password ?? "";
            this.requestQueue = requestQueue;
            this.responseExchange = responseExchange;
            this.responseRoutingKey = responseRoutingKey;

            try
            {
                rabbitMqFactory = new ConnectionFactory() { Uri = new Uri(baseUrl), UserName = username, Password = password, DispatchConsumersAsync = true };
                rabbitMqFactory.AutomaticRecoveryEnabled = true;
                rabbitMqConnection = rabbitMqFactory.CreateConnection();
                rabbitMqChannel = rabbitMqConnection.CreateModel();
                rabbitMqConsumer = new AsyncEventingBasicConsumer(rabbitMqChannel);
                rabbitMqConsumer.Received += async (sender, args) =>
                {
                    await Task.Yield(); // Force async execution.

                    handleRequest(args);

                    // Acknowledge request if response was written successfully.
                    lock (rabbitMqChannel)
                    {
                        try
                        {
                            rabbitMqChannel.BasicAck(deliveryTag: args.DeliveryTag, multiple: false);
                        }
                        catch (Exception ex)
                        {
                            PSScriptInvoker.logError("Unexpected exception while acknowledging request message:\n" + ex.ToString());
                        }

                    }
                };

                rabbitMqChannel.BasicConsume(queue: this.requestQueue,
                                autoAck: false,
                                consumer: rabbitMqConsumer);

                PSScriptInvoker.logInfo("Message consumer successfully started. Now waiting for requests in queue " + this.requestQueue + "...");
            }
            catch (Exception ex)
            {
                string msg = "Unexpected exception while setting up RabbitMQ connection:\n" + ex.ToString();
                PSScriptInvoker.logError(msg);
            }
        }

        private void handleRequest(BasicDeliverEventArgs args)
        {
            try
            {
                IDictionary<string, object>  requestHeaders = args.BasicProperties.Headers;
                string executionId = "";
                string endpoint = "";
                string paramJsonString = "";
                try
                {
                    requestHeaders.TryGetValue("executionId", out object executionIdBytes);
                    requestHeaders.TryGetValue("endpoint", out object endpointBytes);
                    executionId = Encoding.UTF8.GetString((byte[])executionIdBytes);
                    endpoint = Encoding.UTF8.GetString((byte[])endpointBytes);
                    paramJsonString = Encoding.UTF8.GetString(args.Body);
                }
                catch (Exception ex)
                {
                    PSScriptInvoker.logError("Unexpected exception while parsing request segments and query:\n" + ex.ToString());
                    writeResponse(args, "ERROR: URL not valid: " + ex.ToString(), 400);
                    return;
                }

                PSScriptInvoker.logInfo(string.Format("Received RabbitMQ message (deliveryTag: {0}, executionId: {1}, endpoint: {2}):\n{3}", args.DeliveryTag, executionId, endpoint, paramJsonString));

                // Get parameters
                Dictionary<String, String> parameters = JsonConvert.DeserializeObject<Dictionary<String, String>>(paramJsonString);

                // Execute the appropriate script.
                string[] segments = endpoint.Split('/');
                Dictionary<String, String> scriptOutput = scriptExecutor.executePowershellScriptByHttpSegments(segments, parameters);

                // Get output variables
                scriptOutput.TryGetValue("exitCode", out string exitCode);
                scriptOutput.TryGetValue("result", out string result);

                if (exitCode == "0")
                {
                    if (string.IsNullOrEmpty(result))
                    {
                        writeResponse(args, result, 204);
                    }
                    else
                    {
                        writeResponse(args, result, 200);
                    }
                }
                else
                {
                    writeResponse(args, result, 500);
                }
            }
            catch (Exception ex)
            {
                PSScriptInvoker.logError("Unexpected exception while processing message:\n" + ex.ToString());
                writeResponse(args, ex.ToString(), 500);
            }
        }

        private void writeResponse(BasicDeliverEventArgs args, string messageText, int statusCode)
        {
            byte[] messageBytes = Encoding.UTF8.GetBytes(messageText);

            IBasicProperties props = rabbitMqChannel.CreateBasicProperties();
            props.ContentType = "application/json";
            props.DeliveryMode = 2;

            props.Headers = args.BasicProperties.Headers;
            props.Headers.Add("statusCode", statusCode);

            lock (rabbitMqChannel)
            {
                try
                {
                    rabbitMqChannel.BasicPublish(responseExchange, responseRoutingKey, props, messageBytes);
                }
                catch (Exception ex)
                {
                    PSScriptInvoker.logError("Unexpected exception while publishing response message:\n" + ex.ToString());
                }
            }
        }

        public bool isConnected()
        {
            return (rabbitMqConnection != null && rabbitMqConnection.IsOpen && rabbitMqChannel != null);
        }

        public void stopModule()
        {
            rabbitMqConnection.Close();
        }
    }
}
