# Updating existing permissions to include the Delegate flag
Set-MailboxFolderPermission -Identity "MDietz@blackstonepc.com:\Calendar" `
    -User "calendar@blackstonepc.com" `
    -AccessRights Editor `
    -SharingPermissionFlags Delegate