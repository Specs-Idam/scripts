$runAt = Get-Date -Hour 15 -Minute 0 -Second 0
$now = Get-Date

$timeFormat = 'HH'

while($true){

$now = Get-Date
  if ( $runAt.ToString($timeFormat) -eq $now.ToString($timeFormat) ) {

########################################################################


$lettersa = "A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"
$lettersb = "a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z"

$offset = 0
$all_users = @()

rm -Path .\export.csv

foreach($letter1 in $lettersa){

$response = Invoke-RestMethod -Method POST -UseBasicParsing -Uri "https://specsavers.api.identitynow.com/oauth/token?grant_type=client_credentials`&client_id=d64ac2e6-2920-454a-b91b-faa95e2dbbf3`&client_secret=8b8a106005120ca3573acaa9f37f94039f43ac889e2c92a7784c1869fba21140" `
-ContentType 'application/json'

$offset = 0

do{

$body = "{
    `"indices`": [
        `"identities`"
    ],
    `"queryType`": `"SAILPOINT`",
    `"query`": {
        `"query`": `"identityProfile.name:\`"SuccessFactors\`" AND attributes.businessType:Retail AND !attributes.cloudLifecycleState:Inactive AND !attributes.country:\`"Spain\`" AND !attributes.company:Newmedica AND attributes.firstname.exact:$letter1*`"
    },
    `"sort`": [
        `"firstName`"
    ]
}"

    try{
    $searches = Invoke-RestMethod -Method POST -UseBasicParsing -Uri "https://specsavers.api.identitynow.com/v3/search?limit=250&offset=$($offset)" `
    -Headers @{"Authorization" = "Bearer $($response.access_token)" } -ContentType 'application/json' -Body $body}
    catch{break}

    foreach($identity in $searches){$all_users += $identity}

    echo "First $offset $letter1"
    $offset += 250
}while($searches.Count -gt 0)
}

foreach($letter1 in $lettersb){

$response = Invoke-RestMethod -Method POST -UseBasicParsing -Uri "https://specsavers.api.identitynow.com/oauth/token?grant_type=client_credentials`&client_id=d64ac2e6-2920-454a-b91b-faa95e2dbbf3`&client_secret=8b8a106005120ca3573acaa9f37f94039f43ac889e2c92a7784c1869fba21140" `
-ContentType 'application/json'

$offset = 0

do{

$body = "{
    `"indices`": [
        `"identities`"
    ],
    `"queryType`": `"SAILPOINT`",
    `"query`": {
        `"query`": `"identityProfile.name:\`"SuccessFactors\`" AND attributes.businessType:Retail AND !attributes.cloudLifecycleState:Inactive AND !attributes.country:\`"Spain\`" AND !attributes.company:Newmedica AND attributes.firstname.exact:$letter1*`"
    },
    `"sort`": [
        `"firstName`"
    ]
}"

    try{
    $searches = Invoke-RestMethod -Method POST -UseBasicParsing -Uri "https://specsavers.api.identitynow.com/v3/search?limit=250&offset=$($offset)" `
    -Headers @{"Authorization" = "Bearer $($response.access_token)" } -ContentType 'application/json' -Body $body}
    catch{break}

    foreach($identity in $searches){$all_users += $identity}

    echo "Second $offset $letter1"
    $offset += 250
}while($searches.Count -gt 0)
}

$idndata = ".\export.csv"

$fcolumns = "First Name,Last Name,Work E-mail,Employee Number,adUpn,country,nationalIdentifierCompact,sf2PayrollNumber,startDate"
Add-Content -Path $idndata -Value $fcolumns

foreach($user in $all_users){
    
    $string = "$($user.attributes.firstname),$($user.attributes.lastname),$($user.attributes.email),$($user.employeeNumber),$($user.attributes.adUpn),$($user.attributes.country),$($user.attributes.nationalIdentifierCompact),$($user.attributes.sf2PayrollNumber),$($user.attributes.startDate)"
    Add-Content -Path $idndata -Value $string -Encoding UTF8

}

######################################################


$start = Get-Date
$date = Get-Date

$date = $date.toString("ddMMyy")
$filename = ".\report_$($date).csv"

$idn_users = Import-Csv -Path .\export.csv

$final = $filename
$fcolumns = "Username,Work Email,Forename,Surname,Store Names,Store Numbers,Stores Managed,Region,Region Code,Job Title,NI Number,Employee Number,Company Start Date,"
$fcolumns += "Division,Division Code,UK,UK Stores,Primary Store"
Add-Content -Path $final -Value $fcolumns

$start = Get-Date

$creds = [pscredential]::new("SF_IDAM@specSF2PROD", ("Tuesday!4" | ConvertTo-SecureString -AsPlainText -Force))

$json = [collections.arraylist]@()

$allusers = [collections.arraylist]@()

$uri = "https://api5.successfactors.eu/odata/v2/User?`$custom03=Retail (RETAIL)`&custom01=Active`&`$format=json`&`$select=userId,custom02,title,username,email,firstName,lastName,hireDate,empId,country`&`$skiptoken=$($token)"

echo "Pulling SF2 Data..."

do {

    $response = Invoke-RestMethod -Method GET -Credential $creds -Uri $uri `
        -Headers @{"Authorization" = "Basic U0ZfSURBTUBzcGVjU0YyUFJPRDpUdWVzZGF5ITQ=" } -UseBasicParsing -ContentType 'application/json'

    $uri = $response.d.__next

    foreach ($result in $response.d.results) { 

        $json += $result       
    }

}while ($uri)

echo "Sorting SF2 data"

$json = $json | Sort-Object -Property "firstName"

$countries = @("United Kingdom", "Ireland", "Guernsey", "Isle of Man")

$json = $json | Where-Object -FilterScript { $_.country -in $countries }

$index = @($null) * 5000000

echo "Creating data dictionary"

foreach ($row in $json) {
    $key = $row.('empId')

    $data = $index[$key]
    if ($data -is [Collections.ArrayList]) {
        $data.add($row) >$null
    }
    elseif ($data) {
        $index[$key] = [Collections.ArrayList]@($data, $row)
    }
    else {
        $index[$key] = $row
    }

}

echo "Comparing data"

foreach ($idn in $idn_users) {


    $allstores = [collections.arraylist]@()
    $pstorecount = 0
    $nonpstorecount = 0
    $pstore = ""
    $startdate = [datetime]$idn.startDate

    if (($startdate -gt $start.AddDays(1))) { continue }

    $matchSF2 = $index[$idn.'Employee Number']

    if ($matchSF2 -eq $null) {

        $number = $idn.'Employee Number'
        $matchSF2 = Invoke-RestMethod -Method GET -Credential $creds -Uri "https://api5.successfactors.eu/odata/v2/User('$number')?`$format=json`&`$select=userId,custom02,title,username,email,firstName,lastName,hireDate,empId,country" `
            -Headers @{"Authorization" = "Basic U0ZfSURBTUBzcGVjU0YyUFJPRDpUdWVzZGF5ITQ=" } -UseBasicParsing -ContentType 'application/json'
        $matchSF2 = $matchSF2.d
    }
    
    foreach ($SF2 in $matchSF2) {

        if ($SF2.userId.Contains("-")) {
            $pstore = 0
            $nonpstorecount += 1
        }
        elseif (!$SF2.userId.Contains("-")) {
            $pstore = 1
            $pstorecount += 1
        }

        if ($SF2.hireDate -gt $startdate) { $hired = $startdate }else { $hired = $SF2.hireDate }
    
        $entry = [PSCustomObject]@{
            userName  = $idn.adUpn
            email     = $idn.'Work E-mail'
            forename  = $idn.'First Name'
            surname   = $idn.'Last Name'
            store     = $SF2.custom02
            title     = $SF2.title
            NiNo      = $idn.nationalIdentifierCompact
            empNo     = $idn.sf2PayrollNumber
            startDate = $hired
            country   = $idn.country
            userId    = $SF2.userId
            empId     = $SF2.empId
            payroll   = $idn.sf2PayrollNumber
            pstore    = $pstore
        }

        $storeno = $entry.store.Substring($entry.store.LastIndexOf("(") + 1, 4)

        if (($storeno -eq "0018") -or ($storeno -eq "0942") -or ($storeno -eq "0020") -or ($storeno -eq "1494") -or ($storeno -eq "1617") -or ($storeno -eq "0220") -or ($storeno -eq "0025")) {
            if ($entry.pstore -eq 1) { $pstorecount -= 1 }
            elseif ($entry.pstore -eq 0) { $nonpstorecount -= 1 }

        }
        else {
            $allstores.Add($entry)
        }

    }

    if (($nonpstorecount -gt 0) -and ($pstorecount -eq 0)) {
        $pstore = 1
        $allstores = $allstores | Sort-Object -Property startDate
        $allstores[0].pstore = $pstore
    }
    if (($pstorecount -gt 1)) {
        $pstore = 0
        $allstores = $allstores | Sort-Object -Property startDate -Descending

        foreach ($store in $allstores) {
        if($store.pstore -eq 1){$pstorecount -= 1}
        $store.pstore = $pstore
        
        if($pstorecount -eq 1){break;}
    }
    }
    
    foreach ($store in $allstores) {
        $allusers.Add($store)
    }

}

echo "Creating Spreadsheet"

foreach ($entry in $allusers) {

    $userId = $entry.userId
    try {
        $storeno = $entry.store.Substring($entry.store.LastIndexOf("(") + 1, 4)
        $storeno = $storeno -replace '^0+',''
    }
    catch {
        $storeno = ""
    }
    try {
        $storename = $entry.store.Substring(0, $entry.store.LastIndexOf(" "))
    }
    catch {
        $storename = ""
    }
    try {
        $jobtitle = $entry.title.Substring(0, $entry.title.LastIndexOf(" "))
    }
    catch {
        $jobtitle = ""
    }
    
    $username = $entry.username
    $email = $entry.email
    $firstname = $entry.forename
    $lastname = $entry.surname
    $payroll = ""
    $hired = $entry.startDate
    $hired = $hired.toString("dd/MM/yyyy HH:mm")
    $pstore = $entry.pstore

    $NI = $entry.NiNo

    if ($entry.country -eq "Ireland") { $uk = "ROI" }else { $uk = "UK" }
    if ($entry.country -eq "Ireland") { $ukstores = "ROI Store" }else { $ukstores = "UK Store" }
    if ($entry.payroll.Length -eq 7) { $payroll = "00" + $entry.payroll }elseif ($entry.payroll.Length -eq 8) { $payroll = "0" + $entry.payroll }elseif ($entry.payroll.Length -eq 6) { $payroll = "000" + $entry.payroll }else { $payroll = $entry.payroll }
    if ($jobtitle.Contains("Manager")) { $managed = "TRUE" }elseif ($jobtitle.Contains("Director")) { $managed = "TRUE" }elseif ($jobtitle.Contains("Partner")) { $managed = "TRUE" }else { $managed = "FALSE" }
    if (($NI -eq "") -and ($payroll -ne "")) { $NI = $payroll }

    if ($NI -eq "") { $NI = $entry.userName }

    if($username -eq '[no upn]'){continue}
    if(($email -like 'noemail*') -or ($email -like 'noreply*') -or ($email -eq "[no email]")){$email = ""}

    $string = "$($username),$($email),$($firstname),$($lastname),$($storename),$($storeno),$($managed),,,$($jobtitle),$($NI),$($payroll),$($hired),,,$($uk),$($ukstores),$($pstore)"
    Add-Content -Path $final -Value $string

}

& "C:\Users\jonathan.brown\AppData\Local\Programs\WinSCP\WinSCP.com" `
  /command `
    "open sftp://sftp.specsaverslrn:deWmKXYw7agy@sftp.kallidus-suite.com/ -hostkey=`"`"ecdsa-sha2-nistp384 384 Z5GmDQVZ43ayy+nXzS2qGUZ/4z2ytcX9pRIM8/blBz0=`"`" -rawsettings FSProtocol=2" `
    "put $($filename)" `
    "exit"

$start
Get-Date


}
  else {
    if ( $runAt -lt $now ) {
      $runAt.AddDays(1)
    }

  }
  echo $now.ToString($timeFormat)
  Start-Sleep -Seconds 3600
  }
