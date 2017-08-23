using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace PSScriptInvoker
{
    class RabbitMqModule
    {
        private PSScriptExecutor scriptExecutor;

        private string baseUrl;
        private string username;
        private string password;
        private string requestQueueName;
        private string responseQueueName;

        private ConnectionFactory rabbitmqFactory;
        private IConnection rabbitmqConnection;
        private IModel rabbitmqChannel;
        private AsyncEventingBasicConsumer rabbitmqConsumer;

        public RabbitMqModule(PSScriptExecutor scriptExecutor, string baseUrl, string username, string password, string requestQueueName, string responseQueueName)
        {
            this.scriptExecutor = scriptExecutor;
            this.baseUrl = baseUrl;
            this.username = username ?? "";
            this.password = password ?? "";
            this.requestQueueName = requestQueueName;
            this.responseQueueName = responseQueueName;

            try
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

                    PSScriptInvoker.logInfo(string.Format("Received RabbitMQ message (deliveryTag: {0}):\n{1}", args.DeliveryTag, message));

                    handleRequest(message);

                    lock (rabbitmqChannel)
                    {
                        rabbitmqChannel.BasicAck(deliveryTag: args.DeliveryTag, multiple: false);
                    }
                };

                rabbitmqChannel.BasicConsume(queue: requestQueueName,
                                autoAck: false,
                                consumer: rabbitmqConsumer);

                PSScriptInvoker.logInfo("Message consumer successfully started. Now waiting for requests in queue " + requestQueueName + "...");
            }
            catch (Exception ex)
            {
                string msg = "Unexpected exception while setting up RabbitMQ connection:\n" + ex.ToString();
                PSScriptInvoker.logError(msg);
            }
        }

        private void handleRequest(string message)
        {
            try
            {
                // Get parameters
                Dictionary<String, String> parameters = new Dictionary<String, String>();

                // Execute the appropriate script.
                string scriptName = "";
                string scriptPath = "";
                string fullScriptPath = scriptExecutor.getPathToScripts() + scriptPath + scriptName + ".ps1";
                Dictionary<String, String> scriptOutput = scriptExecutor.executePowershellScript(fullScriptPath, parameters);

                // Get output variables
                string exitCode = "";
                string result = "";
                scriptOutput.TryGetValue("exitCode", out exitCode);
                scriptOutput.TryGetValue("result", out result);

                string msg = string.Format("Executed script was: {0}. Exit code: {1}, output:\n{2}", fullScriptPath, exitCode, result);
                PSScriptInvoker.logInfo(msg);

                if (exitCode == "0")
                {
                    if (string.IsNullOrEmpty(result))
                    {
                        writeResponse("");
                        //writeResponse(context.Response, result, 204);
                    }
                    else
                    {
                        writeResponse("");
                        //writeResponse(context.Response, result, 200);
                    }
                }
                else
                {
                    writeResponse("");
                    //writeResponse(context.Response, result, 500);
                }
            }
            catch (Exception ex)
            {
                PSScriptInvoker.logError("Unexpected exception while processing message:\n" + ex.ToString());
                writeResponse(ex.ToString());
                //writeResponse(context.Response, ex.ToString(), 500);
            }
        }

        private void writeResponse(string messageText)
        {

        }

        public bool isConnected()
        {
            return (rabbitmqConnection != null && rabbitmqConnection.IsOpen && rabbitmqChannel != null);
        }

        public void stopModule()
        {
            rabbitmqConnection.Close();
        }
    }
}
