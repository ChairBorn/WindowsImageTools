$attorney = "mdietz@blackstonepc.com"

Get-MailboxFolderStatistics -Identity $attorney -FolderScope Calendar |
  Select-Object Name,FolderType,Identity,ItemsInFolder,FolderPath |
  Format-Table -AutoSize
