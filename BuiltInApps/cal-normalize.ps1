$mbx = "mdietz@blackstonepc.com"

Get-MailboxFolderStatistics -Identity $mbx -FolderScope Calendar |
  Select-Object Name,FolderType,Identity,ItemsInFolder,FolderPath |
  Sort-Object ItemsInFolder -Descending |
  Format-Table -AutoSize
