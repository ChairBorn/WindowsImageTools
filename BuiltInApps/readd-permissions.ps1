$target = "mdietz@blackstonepc.com"
$mgr = "calendar@blackstonepc.com"

# Remove existing entry (use the display name shown if needed)
Remove-MailboxFolderPermission -Identity "$target`:\Calendar" -User $mgr -Confirm:$false -ErrorAction SilentlyContinue
Remove-MailboxFolderPermission -Identity "$target`:\Calendar" -User "Blackstone Calendaring" -Confirm:$false -ErrorAction SilentlyContinue

# Add clean Editor entry
Add-MailboxFolderPermission -Identity "$target`:\Calendar" -User $mgr -AccessRights Editor

# Verify
Get-MailboxFolderPermission -Identity "$target`:\Calendar" |
  Where-Object { $_.User.ToString() -match "Blackstone Calendaring|calendar@blackstonepc\.com" } |
  Format-List *
