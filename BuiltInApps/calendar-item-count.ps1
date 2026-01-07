$mbx = "mdietz@blackstonepc.com"

Get-MailboxFolderStatistics -Identity $mbx -FolderScope Calendar |
  Select-Object Name,FolderType,ItemsInFolder,FolderPath,Identity |
  Sort-Object ItemsInFolder -Descending |
  Format-Table -AutoSize
