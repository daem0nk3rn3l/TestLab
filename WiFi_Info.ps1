#  Get Wi-Fi Name / Password
#
#

(netsh wlan show profiles) | Select-String "\:(.+)$" | %{$name=$_.Matches.Groups[1].Value.Trim(); $_} `
| %{(netsh wlan show profile name="$name" key=clear)} | Select-String "Key Content\W+\:(.+)$" `
| %{$pass=$_.Matches.Groups[1].Value.Trim(); $_} | %{[PSCustomObject]@{ PROFILE_NAME=$name;PASSWORD=$pass }} | Format-Table `
| Out-File "$env:userprofile\OneDrive\Desktop\Wi-Fi_network_passwords.txt"








