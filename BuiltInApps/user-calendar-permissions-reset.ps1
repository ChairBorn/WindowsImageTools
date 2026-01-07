$attorney = "mdietz@blackstonepc.com"
$mgr = "calendar@blackstonepc.com"
$folderId = "$attorney`:\Calendar"

# Remove existing entry (if present)
try {
    Remove-MailboxFolderPermission -Identity $folderId -User $mgr -Confirm:$false -ErrorAction Stop
    "Removed existing permission entry for $mgr on $folderId"
}
catch {
    "Remove failed or entry not present: $($_.Exception.Message)"
}

# Re-add explicitly as Editor
Add-MailboxFolderPermission -Identity $folderId -User $mgr -AccessRights Editor -ErrorAction Stop
"Added Editor permission for $mgr on $folderId"

# Confirm
Get-MailboxFolderPermission $folderId |
Where-Object { $_.User -like "*$mgr*" -or $_.User -eq $mgr } |
Format-List FolderName, User, AccessRights, SharingPermissionFlags
