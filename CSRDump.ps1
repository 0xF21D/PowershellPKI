#requires -modules PSPKI

# CSRDump.ps1
# by Robert Hollingshead
# This is a quick and dirty script using the PSPKI module to dump the subject and san entries for a folder of CSRs into a report object.

$CSRList = Get-ChildItem -Recurse -Path .\ -Filter *.csr

[array]$Report = $Null

ForEach ($CSR in $CSRList) {

    $FileContents = Get-Content $CSR.FullName
    $FileContents = ($FileContents.Replace("-----BEGIN CERTIFICATE REQUEST-----",'')).Replace("-----END CERTIFICATE REQUEST-----",'')
    
    $RawRequest = [convert]::FromBase64String($FileContents)
    $Request = Get-CertificateREquest -RawRequest $RawRequest
    
    $Subject = $Request.Subject.Split('=').Split(',').Split('/')
    $ParentPath = $CSR.Directory -Split ('\\')

    $ReportItem = [PSCustomObject]@{
        Folder = $ParentPath[-1]
        FileName = $CSR.Name
        CommonName = $Null
        Locality = $Null
        State = $Null
        Country = $Null
        OrganizationalUnit = $Null
        Organization = $Null
        emailAddress = $Null
    }

    
    If ($Subject.IndexOf(($Subject | Where-Object{$_ -eq ' O' -or $_ -eq 'O'} | Select -First 1)) -ge 0) {
        $ReportItem.Organization = $Subject[$Subject.IndexOf(($Subject | Where-Object{$_ -eq ' O' -or $_ -eq 'O'} | Select -First 1)) + 1]
    }    
    
    If ($Subject.IndexOf(($Subject | Where-Object{$_ -eq ' OU' -or $_ -eq 'OU'} | Select -First 1)) -ge 0) {
        $ReportItem.OrganizationalUnit = $Subject[$Subject.IndexOf(($Subject | Where-Object{$_ -eq ' OU' -or $_ -eq 'OU'} | Select -First 1)) + 1]
    }
    
    If ($Subject.IndexOf(($Subject | Where-Object{$_ -eq ' CN' -or $_ -eq 'CN'} | Select -First 1)) -ge 0) {
        $ReportItem.CommonName = $Subject[$Subject.IndexOf(($Subject | Where-Object{$_ -eq ' CN' -or $_ -eq 'CN'} | Select -First 1)) + 1]
    }

    If ($Subject.IndexOf(($Subject | Where-Object{$_ -eq ' L' -or $_ -eq 'L'} | Select -First 1)) -ge 0) {
        $ReportItem.Locality = $Subject[$Subject.IndexOf(($Subject | Where-Object{$_ -eq ' L' -or $_ -eq 'L'} | Select -First 1)) + 1]
    }
    
    If ($Subject.IndexOf(($Subject | Where-Object{$_ -eq ' C' -or $_ -eq 'C'} | Select -First 1)) -ge 0) {
        $ReportItem.Country = $Subject[$Subject.IndexOf(($Subject | Where-Object{$_ -eq ' C' -or $_ -eq 'C'} | Select -First 1)) + 1]
    }
    
    If ($Subject.IndexOf(($Subject | Where-Object{$_ -eq ' S' -or $_ -eq 'S'} | Select -First 1)) -ge 0) {
        $ReportItem.State = $Subject[$Subject.IndexOf(($Subject | Where-Object{$_ -eq ' S' -or $_ -eq 'S'} | Select -First 1)) + 1]
    }
    
    If ($Subject.IndexOf(($Subject | Where-Object{$_ -eq ' emailAddress' -or $_ -eq 'emailAddress'} | Select -First 1)) -ge 0) {
        $ReportItem.emailAddress = $Subject[$Subject.IndexOf(($Subject | Where-Object{$_ -eq ' emailAddress' -or $_ -eq 'emailAddress'} | Select -First 1)) + 1]
    }

    $SANCount = ($Request.Extensions[$Request.Extensions.Oid.FriendlyName.IndexOf('Subject Alternative Name')].AlternativeNames.Value).count
    $SANIndex = 0

    

    ForEach ($SAN in $Request.Extensions[$Request.Extensions.Oid.FriendlyName.IndexOf('Subject Alternative Name')].AlternativeNames.Value) {

        $ReportItem | Add-Member -MemberType NoteProperty -Name ("SAN" + $SANIndex) -Value $($SAN)
        $SANIndex++

    }

    $Report += $ReportItem

}


 

