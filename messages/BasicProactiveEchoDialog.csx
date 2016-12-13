#load "Message.csx"

using System;
using System.Threading.Tasks;

using Microsoft.Bot.Builder.Azure;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Microsoft.WindowsAzure.Storage; 
using Microsoft.WindowsAzure.Storage.Queue; 
using Newtonsoft.Json;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;

// For more information about this template visit http://aka.ms/azurebots-csharp-proactive
[Serializable]
public class BasicProactiveEchoDialog : IDialog<object> 
{
    protected ResumptionCookie resumptionCookie = null;

    public Task StartAsync(IDialogContext context)
    {
        context.Wait(MessageReceivedAsync);
        return Task.CompletedTask;
    }

    public virtual async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> argument)
    {
        var message = await argument;

        resumptionCookie = new ResumptionCookie(message);
        
        PromptDialog.Confirm(
            context,
            CreateSiteConfirmAsync,
            "Would you like to create a new site?",
            "Didn't get that!",
            promptStyle: PromptStyle.Auto);
            
            // Create a queue Message
        //    var queueMessage = new Message
        //    {
        //        ResumptionCookie = new ResumptionCookie(message),
        //        Text = message.Text
        //    };

            // write the queue Message to the queue
        //    await AddMessageToQueueAsync(JsonConvert.SerializeObject(queueMessage));

        //    await context.PostAsync($"{this.count++}: You said {queueMessage.Text}. Message added to the queue.");
        //    context.Wait(MessageReceivedAsync);
        //}
    }

    public async Task CreateSiteConfirmAsync(IDialogContext context, IAwaitable<bool> argument)
    {
        var confirm = await argument;
        if (confirm)
        {
            var userName = System.Environment.GetEnvironmentVariable("SPO_U", EnvironmentVariableTarget.Process);
            var password = System.Environment.GetEnvironmentVariable("SPO_P", EnvironmentVariableTarget.Process);
        
            var destinationUrl = "https://rbd3v.sharepoint.com/";
            using (var ctx = new ClientContext(destinationUrl))
            {
             
                ctx.Credentials = new SharePointOnlineCredentials(userName, ConvertToSecureString(password));
                Web web = context.Web;
                ctx.Load(web, w => w.Title, w => w.Language, w => w.Url);
                ctx.ExecuteQueryRetry();
                
                var options = new List<string>(new string[] { "Project", "Event", "Group" });
            
                PromptDialog.Choice(
                context,
                CreateSiteAsync,
                options,
                "Please specify the template you would like to use?",
                "Didn't get that!",
                promptStyle: PromptStyle.Auto);
            }
        }
        else
        {
            await context.PostAsync("Ok, have a nice day");
            context.Wait(MessageReceivedAsync);
        }
    }
    
    public async Task CreateSiteAsync(IDialogContext context, IAwaitable<string> argument)
    {
        var choice = await argument;
        //await context.PostAsync("Message received: " + choice);
        // Create a queue Message
        var queueMessage = new Message
        {
            //ResumptionCookie = new ResumptionCookie(message),
            ResumptionCookie = resumptionCookie,
            Text = choice
        };

        // write the queue Message to the queue
        await AddMessageToQueueAsync(JsonConvert.SerializeObject(queueMessage));

        await context.PostAsync($"Thanks! We are going to create you're site with template '{choice}'. I'm comming back to you when it's ready!");
            
        context.Wait(MessageReceivedAsync);
    }
    
    public static async Task AddMessageToQueueAsync(string message)
    {
        // Retrieve storage account from connection string.
        var storageAccount = CloudStorageAccount.Parse(Utils.GetAppSetting("AzureWebJobsStorage"));

        // Create the queue client.
        var queueClient = storageAccount.CreateCloudQueueClient();

        // Retrieve a reference to a queue.
        var queue = queueClient.GetQueueReference("bot-queue");

        // Create the queue if it doesn't already exist.
        await queue.CreateIfNotExistsAsync();
        
        // Create a message and add it to the queue.
        var queuemessage = new CloudQueueMessage(message);
        await queue.AddMessageAsync(queuemessage);
    }
}

