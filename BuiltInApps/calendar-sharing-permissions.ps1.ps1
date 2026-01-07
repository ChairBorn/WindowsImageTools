$mgr = "calendar@blackstonepc.com"

# Pull all user mailboxes (adjust filters as needed)
$mailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox |
Select-Object -ExpandProperty PrimarySmtpAddress

$results = foreach ($mbx in $mailboxes) {
    $id = "$mbx`:\Calendar"
    try {
        $perm = Get-MailboxFolderPermission $id -ErrorAction Stop |
        Where-Object { $_.User -like "*$mgr*" -or $_.User -eq $mgr }

        if ($perm) {
            foreach ($p in $perm) {
                [pscustomobject]@{
                    Mailbox                = $mbx
                    Folder                 = "Calendar"
                    ManagerEntry           = $p.User.ToString()
                    Rights                 = ($p.AccessRights -join ",")
                    SharingPermissionFlags = $p.SharingPermissionFlags
                }
            }
        }
        else {
            [pscustomobject]@{
                Mailbox                = $mbx
                Folder                 = "Calendar"
                ManagerEntry           = "NOT PRESENT"
                Rights                 = ""
                SharingPermissionFlags = ""
            }
        }
    }
    catch {
        [pscustomobject]@{
            Mailbox                = $mbx
            Folder                 = "Calendar"
            ManagerEntry           = "ERROR"
            Rights                 = ""
            SharingPermissionFlags = $_.Exception.Message
        }
    }
}

# Review in grid, then export if needed
$results | Sort-Object ManagerEntry, Mailbox | Out-GridView

# Optional export
$results | Export-Csv ".\calendar-manager-permissions-audit.csv" -NoTypeInformation
