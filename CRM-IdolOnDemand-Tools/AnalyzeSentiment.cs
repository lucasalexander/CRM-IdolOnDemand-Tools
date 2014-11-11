using System;
using System.Activities;
using System.ServiceModel;
using System.Globalization;
using System.Runtime.Serialization;
using System.IO;
using System.Text;
using System.Net;
using System.Web;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System.Collections.Specialized;

namespace IdolOnDemandTools
{
    public sealed class AnalyzeSentiment : CodeActivity
    {
        [Input("Text Input")]
        public InArgument<String> TextInput { get; set; }

        [Output("Sentiment")]
        public OutArgument<String> Sentiment { get; set; }

        [Output("Score")]
        public OutArgument<Decimal> Score { get; set; }

        //address of the service to which you will post your json messages
        private string _webAddress = "https://api.idolondemand.com/1/api/sync/analyzesentiment/v1";

        //name of your custom workflow activity for tracing/error logging
        private string _activityName = "AnalyzeSentiment";

        //idol ondemand api key
        private string _apiKey = "YOUR_API_KEY_HERE";

        /// <summary>
        /// Executes the workflow activity.
        /// </summary>
        /// <param name="executionContext">The execution context.</param>
        protected override void Execute(CodeActivityContext executionContext)
        {
            // Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            if (tracingService == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve tracing service.");
            }

            tracingService.Trace("Entered " + _activityName + ".Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}",
                executionContext.ActivityInstanceId,
                executionContext.WorkflowInstanceId);

            // Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();

            if (context == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve workflow context.");
            }

            tracingService.Trace(_activityName + ".Execute(), Correlation Id: {0}, Initiating User: {1}",
                context.CorrelationId,
                context.InitiatingUserId);

            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                string inputText = TextInput.Get(executionContext);
                if (inputText != string.Empty)
                {
                    inputText = HtmlTools.StripHTML(inputText);

                    //create the webrequest object and execute it (and post jsonmsg to it)
                    HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(_webAddress);

                    //set request content type so it is treated as a regular form post
                    req.ContentType = "application/x-www-form-urlencoded";

                    //set method to post
                    req.Method = "POST";

                    StringBuilder postData = new StringBuilder();

                    //HttpUtility.UrlEncode
                    //set the apikey request value
                    postData.Append("apikey=" + System.Uri.EscapeDataString(_apiKey) + "&");
                    //postData.Append("apikey=" + _apiKey + "&");

                    //set the text request value
                    postData.Append("text=" + System.Uri.EscapeDataString(inputText));
                    //postData.Append("text=" + inputText);

                    //create a stream
                    byte[] bytes = System.Text.Encoding.ASCII.GetBytes(postData.ToString());
                    req.ContentLength = bytes.Length;
                    System.IO.Stream os = req.GetRequestStream();
                    os.Write(bytes, 0, bytes.Length);
                    os.Close();

                    //get the response
                    System.Net.WebResponse resp = req.GetResponse();

                    //deserialize the response to a SentimentResponse object
                    SentimentResponse myResponse = new SentimentResponse();
                    System.Runtime.Serialization.Json.DataContractJsonSerializer deserializer = new System.Runtime.Serialization.Json.DataContractJsonSerializer(myResponse.GetType());
                    myResponse = deserializer.ReadObject(resp.GetResponseStream()) as SentimentResponse;

                    //set output values from the fields of the deserialzed myjsonresponse object
                    Score.Set(executionContext, myResponse.Aggregate.Score);
                    Sentiment.Set(executionContext, myResponse.Aggregate.Sentiment);
                }
            }

            catch (WebException exception)
            {
                string str = string.Empty;
                if (exception.Response != null)
                {
                    using (StreamReader reader =
                        new StreamReader(exception.Response.GetResponseStream()))
                    {
                        str = reader.ReadToEnd();
                    }
                    exception.Response.Close();
                }
                if (exception.Status == WebExceptionStatus.Timeout)
                {
                    throw new InvalidPluginExecutionException(
                        "The timeout elapsed while attempting to issue the request.", exception);
                }
                throw new InvalidPluginExecutionException(String.Format(CultureInfo.InvariantCulture,
                    "A Web exception ocurred while attempting to issue the request. {0}: {1}",
                    exception.Message, str), exception);
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                tracingService.Trace("Exception: {0}", e.ToString());

                // Handle the exception.
                throw;
            }
            catch (Exception e)
            {
                tracingService.Trace("Exception: {0}", e.ToString());
                throw;
            }

            tracingService.Trace("Exiting " + _activityName + ".Execute(), Correlation Id: {0}", context.CorrelationId);
        }

    }

    //DataContract decoration necessary for serialization/deserialization to work properly
    [DataContract]
    public class SentimentResponse
    {
        //datamember name value indicates name of json field to which data will be serialized/from which data will be deserialized
        [DataMember(Name = "positive")]
        public List<SentimentEntity> Positive { get; set; }

        [DataMember(Name = "negative")]
        public List<SentimentEntity> Negative { get; set; }

        [DataMember(Name = "aggregate")]
        public SentimentAggregate Aggregate { get; set; }

        public SentimentResponse() { }
    }

    //DataContract decoration necessary for serialization/deserialization to work properly
    [DataContract]
    public class SentimentAggregate
    {
        [DataMember(Name = "sentiment")]
        public string Sentiment { get; set; }

        [DataMember(Name = "score")]
        public decimal Score { get; set; }
    }

    //DataContract decoration necessary for serialization/deserialization to work properly
    [DataContract]
    public class SentimentEntity
    {
        [DataMember(Name = "sentiment")]
        public string Sentiment { get; set; }

        [DataMember(Name = "topic")]
        public string Topic { get; set; }

        [DataMember(Name = "score")]
        public decimal Score { get; set; }

        [DataMember(Name = "original_text")]
        public string OriginalText { get; set; }

        [DataMember(Name = "normalized_text")]
        public string NormalizedText { get; set; }

        [DataMember(Name = "original_length")]
        public int OriginalLength { get; set; }

        [DataMember(Name = "normalized_length")]
        public int NormalizedLength { get; set; }

    }
}