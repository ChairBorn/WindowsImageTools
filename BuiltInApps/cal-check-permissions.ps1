$mbx="mdietz@blackstonepc.com"
$mgr="calendar@blackstonepc.com"

$targetFolders = @(
  "/BlackstonePC",
  "/Jonathan M. Genish, Esq.",
  "/Calendar",
  "/Hazel Blackman"
)

foreach ($fp in $targetFolders) {
  $id = "$mbx`:$fp"
  try {
    $p = Get-MailboxFolderPermission -Identity $id -ErrorAction Stop |
         Where-Object { $_.User -like "*$mgr*" -or $_.User -eq $mgr }

    if ($p) {
      foreach ($row in $p) {
        [pscustomobject]@{
          FolderPath = $fp
          Rights     = ($row.AccessRights -join ",")
          Flags      = $row.SharingPermissionFlags
        }
      }
    } else {
      [pscustomobject]@{
        FolderPath = $fp
        Rights     = "NO ENTRY"
        Flags      = ""
      }
    }
  } catch {
    [pscustomobject]@{
      FolderPath = $fp
      Rights     = "ERROR"
      Flags      = $_.Exception.Message
    }
  }
} | Format-Table -AutoSize
