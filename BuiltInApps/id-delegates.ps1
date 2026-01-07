$mailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox |
  Select-Object -ExpandProperty PrimarySmtpAddress

$delegateFindings = foreach ($mbx in $mailboxes) {
  try {
    $perms = Get-MailboxFolderPermission "$mbx`:\Calendar" -ErrorAction Stop |
             Where-Object { $_.SharingPermissionFlags -match "Delegate" }

    foreach ($p in $perms) {
      [pscustomobject]@{
        Mailbox  = $mbx
        Trustee  = $p.User.ToString()
        Rights   = ($p.AccessRights -join ",")
        Flags    = $p.SharingPermissionFlags
      }
    }
  } catch { }
}

$delegateFindings | Format-Table -AutoSize
$delegateFindings | Export-Csv .\DelegateCalendarEntries.csv -NoTypeInformation
