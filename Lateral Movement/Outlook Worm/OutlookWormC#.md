# Lateral Movement: Outlook Worm - C#

Lateral Movement is an important aspect of any offensive campaign. In this article I'm going to show you how to build an Outlook Worm 
in C# for use during adversarial emulation campaigns. Before we begin, please make sure you have Visual Studio installed.

*Note: This method __does not__ require the victim's Outlook username or password. Furthermore, while this tutorial will explain how to go about retrieving victim email addresses and sending emails to them, I __will not__ be covering the actual plannning or generation of a phishing email template.*

Open Visual Studio, and create a new .Net Core Console Application (C#). Once Visual Studio loads your new project, right click on the project name
in the Solution Explorer side menu and click "Manage Nuget Packages"

![Click on Manage Nuget Packlages](https://raw.githubusercontent.com/bcdannyboy/Writing/master/Lateral%20Movement/Outlook%20Worm/managenugetpackages.png)

Next, switch to the "Browse" tab and search for _Microsoft.Office.Interop.Outlook_ click on it and install.

![Microsoft.Office.Interop.Outlook](https://raw.githubusercontent.com/bcdannyboy/Writing/master/Lateral%20Movement/Outlook%20Worm/MicrosoftOfficeInteropOutlook.png)

Now that we have everything set up, we can switch back to out development window (Program.cs) and begin building the worm. Since the purpose of our worm is to spread a malicious email to all of our victim's contacts, let's first start by gathering up target emails from the Outlook sent box and the user's contact list. 

*Note: I've chosen to gather emails from the user's sent box because not everyone adds people to their contacts and most people receive spam emails that we don't want to send our payload to. At least in the sent box we know that the emails there are to real people. Also, since this tutorial is aimed toward those looking to execute an adversarial emulation (Red Team) campaign, I'm going to assume your victims are going to be domain joined.*

lets go ahead and create a function called *GetEmailsFromVictim* that's of return type *List<string>*. and call it into a *List<String>* variable our *Main* function. *Note: Make sure to import System.Collections.Generic so that the List will work.*
  
To gather recipient email addresses, we need to connect to the outlook application, a MAPI namespace within it, and then the folder we want to enumerate. See the screenshot below:

![Get victim's sent box recipients and contacts](https://raw.githubusercontent.com/bcdannyboy/Writing/master/Lateral%20Movement/Outlook%20Worm/GetEmailsFromVictim.PNG)

Now that we have all of our recipients, all we have to do is send them our phishing email and malicious attachment. See the screenshot below:

![Send Emails through Outlook](https://raw.githubusercontent.com/bcdannyboy/Writing/master/Lateral%20Movement/Outlook%20Worm/SendOutlookEmail.PNG)

Now, let's take a look at what our program looks like behind-the-scenes using IDA.

If we look in our *Main()* function in IDA, we can see that the program calls *GetEmailsFromVictim()* as a *List<String>* before initiating the ForEach loop with a *List.GetEnumerator()* call and looping through each victim email address using *Enumerator.MoveNext()*. During the loop, the program gets the current item in the list using *List.get_Current()* and then passes it to the *SendOutlookEmail()* function as the *string address* parameter. Finally, the program disposes of the Enumerator and prints "Done".

![Main() function in IDA](https://raw.githubusercontent.com/bcdannyboy/Writing/master/Lateral%20Movement/Outlook%20Worm/mainida.PNG)

Our GetEmailsFromVictim() function is a bit more complicated:

![GetEmailsFromVictim() call graph zoomed out](https://raw.githubusercontent.com/bcdannyboy/Writing/master/Lateral%20Movement/Outlook%20Worm/complicatedGetVictimEmails.PNG)

Let's take a look at the start of the *GetEmailsFromVictim()* function.

![Start of GetEmailsFromVictim() function](https://raw.githubusercontent.com/bcdannyboy/Writing/master/Lateral%20Movement/Outlook%20Worm/getemailsstart.PNG)

Hete we cam see the creation of the 2 lists as well as the creation of the Outlook object through the *GetTypeFromCLSID* method to get Outlook COM class and the *CreateInstance()* method to create an outlook instance using it's COM class object. Then, we use *GetNamespace()* to get the MAPI namespace, followed by two *GetDefaultFolder()* and *get_Items()* functions to get the Sent Folder and Contact Folder items.

Then, the program moves through it's email retrieval process using the following steps:

![GetEmailsFromVictim() email retrieval process steps](https://raw.githubusercontent.com/bcdannyboy/Writing/master/Lateral%20Movement/Outlook%20Worm/getemailssteps.png)

1. The SentItems loop is initiated.
2. RecipientItem is retrieved from SentItems as a *Outlook.MailItem* object using the *get_Item()* method
3. The email's Recipients are retrieved using the *get_Recipients* method and then the program loops through all the recipients.
4. The current recipient address is retrieved using the *get_AddressEntry()*, *GetExchangeUser()*, and *GetPrimarySMTPAddress()* methods, before being added to our *List<String> GatheredEmailsWithDuplicates* and returning to the start of the loop.
5. After all of the recipients in the sentbox are gathered, we move to the ContactItems loop.
6. We loop through and get a contact using the *get_Item()* method and then get the contact's primary email address using the *get_Email1Address()* method, before adding it to our *List<String> GatheredEmailsWithDuplicates* and returning to the start of the loop.
7. We then start our duplicate removal loop.
8. We get the current address in the loop using the *get_Current()* method, then we create a *Mail.MailAddress()* object with our retrieved address and confirm that the formatted address is the same as the one we passed it using the *get_Address()* and *op_Equality()* functions to retrieve the formatted address and compare the two respectively.
9. We then check if our final list already has the current address (to avoid duplicates) using the  *List.Contains()* method
10. Finally, if the final list does not contain our current address, we add it to the list and return to the start of the loop.

Once our program collects all the available email addresses, we send each of them one by one to the *SendOutlookEmail()* function and begin spreading our payload.

![SendOutlookEmail() function in IDA](https://raw.githubusercontent.com/bcdannyboy/Writing/master/Lateral%20Movement/Outlook%20Worm/SendOutlookEmailIDA.PNG)

In this function, we first create a new Outlook instance through the Outlook COM class object using the *GetTypeFromCLSID()* and *CreateInstance()* functions. Then we crate a new Mail Item using the *CreateItem()* function. After that, we start setting up our email address by creating / setting opur subject and body, adding our victim's email address to the Recipient list, and having the payload add itself as an attachment to the email. Finally, we resolve our recipients using the *Recipients.ResolveAll()* function and send our email using the *MailItem.Send()* function.

To continue with our analysis, if we add the application to Any.run (https://app.any.run/tasks/734b2eb3-4417-4414-a849-e1a2ad5a9c2c) we see that the application executes OUTLOOK.exe via COM and only interacted with the network to connect to Outlook's configuration site (config.messenger.msn.com).

![Any.Run results](https://raw.githubusercontent.com/bcdannyboy/Writing/master/Lateral%20Movement/Outlook%20Worm/anyrun.PNG)

Interestingly, if we upload the file to VirusTotal, it's identified as a 1/67 with Cylance marking it as "Unsafe"

![VirusTotal 1/67 because of Cylance](https://raw.githubusercontent.com/bcdannyboy/Writing/master/Lateral%20Movement/Outlook%20Worm/virustotal.PNG)

I have uploaded the source code to this tutorial in this github directory. 

Thanks! I hope you learned something!
