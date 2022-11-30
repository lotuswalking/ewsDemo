// See https://aka.ms/new-console-template for more information

using System;
using System.Security;
using Microsoft.Exchange.WebServices.Data;


String ewsUrl = "https://mail.lenovo.com/ews/exchange.asmx";
String userName, password, mailbox;
Console.WriteLine("Please Input User name(xxx@xxx.xxx):");
userName = Console.ReadLine();
Console.WriteLine("Please Input Password For User:"+userName);
password = getPassword();

ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2016);
service.Credentials = new WebCredentials(userName,password);
service.TraceEnabled = false;
service.TraceFlags = TraceFlags.DebugMessage;
service.Url = new Uri(ewsUrl);
//listEmail(service, new FolderId(WellKnownFolderName.Inbox, userName));
sendEmail(service, new FolderId(WellKnownFolderName.SentItems, userName));

//Console.ReadLine();
static void sendEmail(ExchangeService service, FolderId folder)
{
    try
    {
        EmailMessage email = new EmailMessage(service);
        email.ToRecipients.Add("lijy1@lenovo.com");
        email.Subject = "HelloWorld";
        email.Body = new MessageBody("This is the first email I've sent by using the EWS Managed API");
        //email.Save(folder);
        email.SendAndSaveCopy();
        Console.WriteLine("***Email Send Successed*****");
    }
    catch (Exception e)
    {
        Console.WriteLine("*****Email Send Failed be Below error*********");
        Console.WriteLine(e.ToString());
    }
}
static void listEmail(ExchangeService service, FolderId folder)
{
    try
    {
        //SearchFilter sf = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
        ItemView view = new ItemView(10);
        view.Traversal = ItemTraversal.Shallow;
        view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);
        //FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, sf, view);
        FindItemsResults<Item> findResults = service.FindItems(folder, view);
        foreach (Item item in findResults)
        {
            Console.WriteLine(item.DateTimeReceived.ToString()+":"+item.Subject);
        }
        Console.WriteLine("Email List Successed!");
    }
    catch (Exception e)
    {

        Console.WriteLine(e.ToString());
    }
    //Console.WriteLine(findResults.Count());
}

static String getPassword()
{
    var pass = string.Empty;
    ConsoleKey key;
    do
    {
        var keyInfo = Console.ReadKey(intercept: true);
        key = keyInfo.Key;

        if (key == ConsoleKey.Backspace && pass.Length > 0)
        {
            Console.Write("\b \b");
            pass = pass[0..^1];
        }
        else if (!char.IsControl(keyInfo.KeyChar))
        {
            Console.Write("*");
            pass += keyInfo.KeyChar;
        }
    } while (key != ConsoleKey.Enter);
    Console.WriteLine();
    return pass;
}
