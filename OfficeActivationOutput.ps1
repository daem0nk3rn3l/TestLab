# Check license activation
#
# Office 32bit:   CD "%SystemDrive%\Program Files\Microsoft Office\Office16"
# Check office activation:   cscript ospp.vbs /dstatus
# 



enum Licensestatus{
Unlicensed = 0
Licensed = 1
Out_Of_Box_Grace_Period = 2
Out_Of_Tolerance_Grace_Period = 3
Non_Genuine_Grace_Period = 4
Notification = 5
Extended_Grace = 6
}

$result = Get-CimInstance -ClassName SoftwareLicensingProduct | where {$_.name -like "*office*"}| select Name, ApplicationId, @{N='LicenseStatus'; E={[LicenseStatus]$_.LicenseStatus}}
$result|Out-GridView


# ------------------------------------------------------------------------------------------------
# Remove the registry keys
# Remove-Item –Path “HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\OEM” –Recurse
# Remove-Item –Path “HKLM:\ SOFTWARE\Microsoft\Office\16.0\Common\OEM” –Recurse


