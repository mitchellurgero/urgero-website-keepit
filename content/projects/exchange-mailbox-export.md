---
title: "Exchange Mailbox Export"
date: 2019-04-01T11:49:30-05:00
draft: false
categories: ["Uncategorized"]
---

[Project Page](https://gist.github.com/mitchellurgero/45b52fd044a67d9a768b507a7fa78a6d)

After much searching around and trying to get 2013 / 2016 exchange to properly export to PST any mailbox and items that included my search terms.

On the project page you will find a PowerShell script that is easily configurable. This script will allow you to specify user mailbox's to search, where to put the found items, and where to export the new PST to. The PST will contain the emails, in a folder, for each mailbox. 

In the script, you will find 5 variables. These variables will assist you in getting that sweet sweet PST file for viewing search results.

- $idents
    - Idents variable is where you will put the username's for each mailbox. For exchange 2013/2016 that will be formatted like: `"user1","user2","user3"`

- $targetMailbox
    - TargetMailbox is the mailbox to save the search results (found mail items) to. This can be your mailbox, or a new one you can create in the Exchange Control Panel.
	
- $folderName
    - FolderName is the folder you want to save to in the `$targetMailbox`. I enjoy formatting this like `Search 06/06/2018`.
	
- $searchTerm
    - SearchTerm is the search phrase you are looking for in emails. This is compatible with [Outlook search syntax](https://support.office.com/en-us/article/learn-to-narrow-your-search-criteria-for-better-searches-in-outlook-d824d1e9-a255-4c8a-8553-276fb895a8da)
	
- $resultPath
    - ResultPath is the UNC path to save the PST to. This is formatted as `\\SERVERNAME\SHARENAME\FILENAME.PST` It *MUST* be a UNC Fileshare.
	
Once you have the above variables set, you will get something that looks like:

{{< highlight powershell >}}
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;

$idents        = "user1","user2","user3"
$targetMailbox = "user4"
$folderName    = "Search 01/01/2018"
$searchTerm    = "Phrase to search for"
$resultPath    = "\\SRV-EX1\Private\PSTOUT\search01012018.pst"

foreach($i in $idents){
    Write-Host "Searching $i"
    search-mailbox -Identity "$i" -searchquery "$searchTerm" -targetmailbox "$targetMailbox" -TargetFolder "$folderName" -loglevel full
}

New-MailboxExportRequest -Mailbox "$targetMailbox" -IncludeFolders "$folderName" -FilePath "$resultPath"
{{< / highlight >}}

Then just run the script, and you will see a new PST in the ResultPath you set!