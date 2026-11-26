Connect-ExchangeOnline -DelegatedOrganization <Domain>.onmicrosoft.com -ShowBanner:$false

Get-Mailbox -ResultSize Unlimited |
    Where-Object { $_.DeliverToMailboxAndForward -eq $true } |
    Select-Object DisplayName, UserPrincipalName, ForwardingSMTPAddress, ForwardingAddress |
    Export-Csv -Path ".\ForwardingMailboxes.csv" -NoTypeInformation -Encoding UTF8

Write-Output "Export complete."