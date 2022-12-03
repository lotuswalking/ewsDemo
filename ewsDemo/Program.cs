// See https://aka.ms/new-console-template for more information

using System;
using System.Numerics;
using System.Runtime.InteropServices;
using System.Security;
using Microsoft.Exchange.WebServices.Data;



//String ewsUrl = "https://mail.lenovo.com/ews/exchange.asmx";
//String ewsUrlOL = "https://outlook.office365.com/EWS/Exchange.asmx ";
String userName, password, mailbox,mailhost;
//Console.WriteLine("Where your want to Connect to? For Office 365 you could input:(outlook.office365.com)");
mailhost = ReadLine.Read("Where your want to Connect to? For Office 365 you could input:outlook.office365.com\n","outlook.office365.com");
String ewsUrl = String.Format("https://{0}/ews/exchange.asmx",mailhost);
Console.WriteLine("Ews URL is: "+ewsUrl);
String actionFlag = ReadLine.Read("What's do you want to do? 1 For list 10 email(default), 2 for send test email to your self: 1\n", "1");
userName = ReadLine.Read("Please Input User name(xxx@xxx.xxx):");
Console.WriteLine("Please Input Password For User:"+userName);
password = getPassword();

ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2016);
service.Credentials = new WebCredentials(userName,password);
service.TraceEnabled = false;
service.TraceFlags = TraceFlags.DebugMessage;
service.Url = new Uri(ewsUrl);
if(actionFlag.Trim().Equals("1")) {
    listEmail(service, new FolderId(WellKnownFolderName.Inbox, userName));
}
else if(actionFlag.Trim().Equals("2")) {
    sendEmail(service, new FolderId(WellKnownFolderName.SentItems, userName),userName);
}else
{
    Console.WriteLine("No Action Defined");
}

ReadLine.Read("presss enter to Quit!");

static void sendEmail(ExchangeService service, FolderId folder,String recpt)
{
    try
    {
        EmailMessage email = new EmailMessage(service);
        //String recpt = "youremailAddress@gmail.com";
        email.ToRecipients.Add(recpt);
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
