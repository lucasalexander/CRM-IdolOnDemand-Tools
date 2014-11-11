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
    public sealed class IndexEmail : CodeActivity
    {
        [Input("Content")]
        public InArgument<String> Content { get; set; }

        [Input("Subject")]
        public InArgument<String> Subject { get; set; }

        [Input("Email")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> Email { get; set; }

        //address of the service to which you will post your json messages
        private string _webAddress = "https://api.idolondemand.com/1/api/sync/addtotextindex/v1";

        //address of the service to which you will post your json messages
        private string _indexName = "YOUR_INDEX_NAME";

        //name of your custom workflow activity for tracing/error logging
        private string _activityName = "IndexEmail";

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
                string inputText = Content.Get(executionContext);
                if (inputText != string.Empty)
                {
                    inputText = HtmlTools.StripHTML(inputText);

                    IndexDocument myDoc = new IndexDocument
                    {
                     Content = inputText,
                      Reference = (Email.Get(executionContext)).Id.ToString(),
                      Subject = Subject.Get(executionContext),
                      Title = Subject.Get(executionContext)
                    };

                    DocumentWrapper myWrapper = new DocumentWrapper();
                    myWrapper.Document = new List<IndexDocument>();
                    myWrapper.Document.Add(myDoc);

                    //serialize the myjsonrequest to json
                    System.Runtime.Serialization.Json.DataContractJsonSerializer serializer = new System.Runtime.Serialization.Json.DataContractJsonSerializer(myWrapper.GetType());
                    MemoryStream ms = new MemoryStream();
                    serializer.WriteObject(ms, myWrapper);
                    string jsonMsg = Encoding.Default.GetString(ms.ToArray());

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

                    //set the json request value
                    postData.Append("json=" + jsonMsg + "&");

                    //set the index name request value
                    postData.Append("index=" + _indexName);

                    //create a stream
                    byte[] bytes = System.Text.Encoding.ASCII.GetBytes(postData.ToString());
                    req.ContentLength = bytes.Length;
                    System.IO.Stream os = req.GetRequestStream();
                    os.Write(bytes, 0, bytes.Length);
                    os.Close();

                    //get the response
                    System.Net.WebResponse resp = req.GetResponse();

                    //deserialize the response to a ResponseBody object
                    ResponseBody myResponse = new ResponseBody();
                    System.Runtime.Serialization.Json.DataContractJsonSerializer deserializer = new System.Runtime.Serialization.Json.DataContractJsonSerializer(myResponse.GetType());
                    myResponse = deserializer.ReadObject(resp.GetResponseStream()) as ResponseBody;

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
    public class DocumentWrapper
    {
        //datamember name value indicates name of json field to which data will be serialized/from which data will be deserialized
        [DataMember(Name = "document")]
        public List<IndexDocument> Document { get; set; }

    }

    //DataContract decoration necessary for serialization/deserialization to work properly
    [DataContract]
    public class IndexDocument
    {
        [DataMember(Name = "title")]
        public string Title { get; set; }

        [DataMember(Name = "reference")]
        public string Reference { get; set; }

        [DataMember(Name = "subject")]
        public string Subject { get; set; }

        [DataMember(Name = "content")]
        public string Content { get; set; }

    }

    //DataContract decoration necessary for serialization/deserialization to work properly
    [DataContract]
    public class ResponseBody
    {
        [DataMember(Name = "index")]
        public string Index { get; set; }

        [DataMember(Name = "references")]
        public List<ResponseReference> References { get; set; }

    }

    //DataContract decoration necessary for serialization/deserialization to work properly
    [DataContract]
    public class ResponseReference
    {
        [DataMember(Name = "reference")]
        public string Reference { get; set; }

        [DataMember(Name = "id")]
        public int Id { get; set; }
    }

}