#  Get Wi-Fi Name / Password
#  https://www.elevenforum.com/t/find-wi-fi-network-security-key-password-in-windows-11.4475/

(netsh wlan show profiles) | Select-String "\:(.+)$" | %{$name=$_.Matches.Groups[1].Value.Trim(); $_} `
| %{(netsh wlan show profile name="$name" key=clear)} | Select-String "Key Content\W+\:(.+)$" `
| %{$pass=$_.Matches.Groups[1].Value.Trim(); $_} | %{[PSCustomObject]@{ PROFILE_NAME=$name;PASSWORD=$pass }} | Format-Table `
| Out-File "$env:userprofile\OneDrive\Desktop\Wi-Fi_network_passwords.txt"
