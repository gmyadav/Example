using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Rest;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;

namespace FunctionApp1;

public class AzureDataverse
{
    private readonly ILogger<AzureDataverse> _logger;

    public AzureDataverse(ILogger<AzureDataverse> logger)
    {
        _logger = logger;
    }

    [Function("CreateContact")]
    public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequest req)
    {
        string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
        var contactData = JsonConvert.DeserializeObject<ContactModel>(requestBody);


        _logger.LogInformation("C# HTTP trigger function processed a request.");
        var SecretValue = Environment.GetEnvironmentVariable("SecretValue");
        var AppID = Environment.GetEnvironmentVariable("AppID");
        var InstanceUri = Environment.GetEnvironmentVariable("InstanceUri");

        string ConnectionStr = $@"AuthType=ClientSecret;
                        SkipDiscovery=true;url={InstanceUri};
                        Secret={SecretValue};
                        ClientId={AppID};
                        RequireNewInstance=true";

        _logger.LogInformation(ConnectionStr);
        var service = new ServiceClient(ConnectionStr);
        Guid contactId = new Guid();
        if (service.IsReady)

        {
            switch (req.Method.ToUpper())
            {
                case "POST":
                    var newContact = new Entity("contact");
                    newContact["firstname"] = contactData?.FirstName;
                    newContact["lastname"] = contactData?.LastName;
                    newContact["emailaddress1"] = contactData?.Email;
                    Guid createdId = service.Create(newContact);
                    return new OkObjectResult(new { Message = "Contact created", Id = createdId });

                case "GET":
                    var query = new QueryExpression("contact")
                    {
                        ColumnSet = new ColumnSet("firstname", "lastname", "emailaddress1"),
                        TopCount = 10
                    };
                    var results = service.RetrieveMultiple(query);
                    var contacts = new List<object>();
                    foreach (var contact in results.Entities)
                    {
                        contacts.Add(new
                        {
                            Id = contact.Id,
                            FirstName = contact.GetAttributeValue<string>("firstname"),
                            LastName = contact.GetAttributeValue<string>("lastname"),
                            Email = contact.GetAttributeValue<string>("emailaddress1")
                        });
                    }
                    return new OkObjectResult(contacts);

                case "PUT":
                    if (Guid.TryParse(contactData?.Id, out Guid updateId))
                    {
                        var updateContact = new Entity("contact") { Id = updateId };
                        updateContact["firstname"] = contactData.FirstName;
                        updateContact["lastname"] = contactData.LastName;
                        updateContact["emailaddress1"] = contactData.Email;
                        service.Update(updateContact);
                        return new OkObjectResult(new { Message = "Contact updated", Id = updateId });
                    }
                    return new BadRequestObjectResult("Invalid or missing contact ID");

                case "DELETE":
                    if (Guid.TryParse(contactData?.Id, out Guid deleteId))
                    {
                        service.Delete("contact", deleteId);
                        return new OkObjectResult(new { Message = "Contact deleted", Id = deleteId });
                    }
                    return new BadRequestObjectResult("Invalid or missing contact ID");

                default:
                    return new BadRequestObjectResult("Unsupported HTTP method");
            }


        }

        return new OkObjectResult("Welcome to Azure Functions! COntact Created  " + contactId);
    }


}