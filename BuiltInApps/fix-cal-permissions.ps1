Add-MailboxFolderPermission -Identity "mdietz@blackstonepc.com:\Calendar" `
  -User calendar@blackstonepc.com `
  -AccessRights Editor

Get-MailboxFolderPermission -Identity "mdietz@blackstonepc.com:\Calendar" |
  Where-Object { $_.User -like "*calendar@blackstonepc.com*" -or $_.User -eq "calendar@blackstonepc.com" } |
  Format-List FolderName,User,AccessRights,SharingPermissionFlags
