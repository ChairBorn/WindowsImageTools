$mbx = "mdietz@blackstonepc.com"
$mgr = "calendar@blackstonepc.com"

$f = Get-MailboxFolderStatistics -Identity $mbx -FolderScope Calendar |
     Where-Object { $_.Name -eq "BlackstonePC" }

# Use FolderId, not the friendly folder name
$id = "$mbx`:\$($f.FolderId)"

Get-MailboxFolderPermission -Identity $id |
  Where-Object { $_.User -like "*$mgr*" -or $_.User -eq $mgr } |
  Format-List FolderName,User,AccessRights,SharingPermissionFlags
