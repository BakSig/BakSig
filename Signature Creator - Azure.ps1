$username= "ADMIN@DOMAIN.COM"
$password = ConvertTo-SecureString "ADMINPASS" -AsPlainText -Force
$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $password
import-module azuread
Connect-AzureAD -Credential $cred


$user = whoami /upn


$folderLocation = Join-Path -Path $Env:appdata -ChildPath 'Microsoft\signatures'
$filename = 'Bakerly Signature'
$file  = Join-Path -Path $folderLocation -ChildPath $filename


if (-not (Test-Path -Path $folderLocation)) {
  try {
      New-Item -ItemType directory -Path $folderLocation
  } catch {
      Write-Host "Error: Unable to create the signatures folder. Details: $($_.Exception.Message)"
      exit
  }
}


$logo = 'https://github.com/BakSig/BakSig/blob/main/Non%20Animated.jpg?raw=true' 



$disp = get-azureaduser -ObjectId $user | select displayname -Expandproperty displayname
$job = get-azureaduser -ObjectId $user| select jobtitle -ExpandProperty jobtitle
$str =get-azureaduser -ObjectId $user| select streetaddress -ExpandProperty streetaddress
$tele = get-azureaduser -ObjectId $user| select Telephonenumber -ExpandProperty Telephonenumber
$Ema = get-azureaduser -ObjectId $user| select UserPrincipalName -ExpandProperty Userprincipalname
$website = "Bakerly.com"

$displayname = $disp
$jobtitle = $job
$street =$str
$telephone = $tele
$Email = $ema



# Building Style Sheet
$style = 
@"
<style>
p, table, td, tr, a, span { 
    font-family: Trebuchet MS
    font-size:  10pt;
    color: #00B6EE;
}

span.blue
{
    color: #00B6EE;
}

table {
    margin: 0;
    padding: 0;
}

a { 
text-decoration: none;
}

hr {
border: none;
height: 1px;
background-color: #28b8ce;
color: #28b8ce;
width: 250px;
}

table.main {
    border-top: 1px solid #00B6EE;
    border-left: 1px solid #00B6EE;
    border-bottom: 1px solid #00B6EE;
    border-right: 1px solid #00B6EE;
}
</style>
"@

# Building HTML
$signature = 
@"

<p>
  <table class='main'>
      <tr>
          <td style='padding-Right: 10px;'>$(if($logo){"<img src='$logo' width='360' height='216' />"})</td>
          <td>
              <table>
              <tr><td colspan='2' style='padding-bottom: 5px;'>
                    $(if($displayName){"<span><b><font size='+2'>"+$displayName+"</b></span><br /></font>"})
                    $(if($jobTitle){"<span><b><font color=#053167>"+$jobTitle+"</span><br /><br />"})                                 
                  </td></tr> 
                  $(if($street){"<tr><td><b>A: </td><td><a href='tel:$street'>$($street)</a></td></tr>"})                 
                  $(if($telephone){"<tr><td><b>P: </td><td><a href='tel:$telephone'>$($telephone)</a></td></tr>"})
                  $(if($mobileNumber){"<tr><td><b>M: </td><td><a href='tel:$mobileNumber'>$($mobileNumber)</a></td></tr>"})
                  $(if($email){"<tr><td><b>E: </td><td><a href='mailto:$email'>$($email)</a></td></tr>"})
                  $(if($website){"<tr><td><b>W: </td><td><a href='https://$website'>$($website)</a></td></tr>"})
                  <tr><td><b> </td><td><a href='1'>$($1)</a></td></tr>

                  <a href="https://www.facebook.com/bakerlyusa/"><img src="https://github.com/BakSig/BakSig/blob/main/Facebook.png?raw=true" alt="Facebook" style="width:30px;height:40px;"></a>
                  <a href="https://www.instagram.com/bakerlyusa/"><img src="https://github.com/BakSig/BakSig/blob/main/Instagram.png?raw=true" alt="Instagram" style="width:30px;height:40px;"></a>
                  <a href="https://twitter.com/bakerlyUSA"><img src="https://github.com/BakSig/BakSig/blob/main/Twitter.png?raw=true" alt="Twitter" style="width:30px;height:40px;"></a>
                  <a href="https://www.linkedin.com/company/bakerly"><img src="https://github.com/BakSig/BakSig/blob/main/Linkedin.png?raw=true" alt="LinkedIn" style="width:30px;height:40px;"></a>

              </table>
          </td>
      </tr>
  </table>
</p>
<br />
"@

# Save the HTML to the signature file
try {
  $style + $signature | Out-File -FilePath "$file.htm" -Encoding ascii
} catch {
  Write-Host "Error: Unable to save the HTML signature file. Details: $($_.Exception.Message)"
  exit
}



try {
  $signature | out-file "$file.txt" -encoding ascii
} catch {
  Write-Host "Error: Unable to save the text signature file. Details: $($_.Exception.Message)"
  exit
}

# SetS the regkeys for Outlook 
if (test-path "HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General") 
{  
    get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name NewSignature -value $filename -propertytype string -force
    get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\MailSettings | new-Itemproperty -name ReplySignature -value $filename -propertytype string -force
    Remove-ItemProperty -Path HKCU:\\Software\\Microsoft\\Office\\16.0\\Outlook\\Setup -Name "First-Run" -ErrorAction silentlycontinue
}

