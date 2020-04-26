<#
.SYNOPSIS
  Email users who has a password expring soon
.DESCRIPTION
  Use this with window task scheduler to emit a password expiring soon email
.NOTES
  Version:       0.1
  Author:        Jeff Guymon <jguymon@gmail.com>
  Creation Date: 2020-04-15
#>


param(
  [String]$recipients,
  [String]$smtpHost = "localhost",
  [String]$sender = "admin@localhost"
)


function Get-ADPassworExpiry {
  Get-ADUser `
    -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False -and mail -like '*'} `
	-Properties `
	  "DisplayName", `
	  "msDS-UserPasswordExpiryTimeComputed", `
	  "EmailAddress" `
  | Select-Object `
    -Property `
	  "DisplayName", `
	  "EmailAddress", `
	  @{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}
}


function Get-ADPassworExpirySoon {
  $timeSpanSoon = (Get-Date) + (New-TimeSpan -Days 7)
  Get-ADPassworExpiry | Where-Object {$_.ExpiryDate -lt $timeSpanSoon -and $_.ExpiryDate -gt (Get-Date)}
}


function Send-ADPassworExpirySoon {
  Get-ADPassworExpirySoon | ForEach-Object {
    $subject = "Your Active Directory ({0}) password is expiring soon" -f $env:computername
    $body = @'
	  Hello {0},<br/>
	  This email is to let you know your Active Directory password for {2} is going to expire on {1}.
	  We recommend you update your password soon. 
'@ -f $_.DisplayName,$_.ExpiryDate,$env:computername
    $hostMessage = "sending email to {0}" -f $_.EmailAddress
	
	Write-Host $hostMessage
	
	Send-MailMessage `
	  -SmtpServer $smtpHost `
	  -To $_.EmailAddress `
	  -From $sender `
	  -Subject $subject `
	  -BodyAsHtml `
	  -body $body
    
  }
}

Send-ADPassworExpirySoon