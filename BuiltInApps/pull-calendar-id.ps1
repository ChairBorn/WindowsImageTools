$mbx = "mdietz@blackstonepc.com"

Get-MailboxFolderStatistics -Identity $mbx -FolderScope Calendar |
  Select-Object Name,FolderType,FolderPath,FolderId,Identity |
  Format-Table -AutoSize
