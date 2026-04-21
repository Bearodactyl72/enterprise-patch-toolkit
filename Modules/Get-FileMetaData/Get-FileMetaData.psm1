# DOTS formatting comment

function Get-FileMetaData {
    <#
    .SYNOPSIS
    Gets metadata information from file providing similar output to what Explorer shows when viewing file
    .DESCRIPTION
    Small function that gets metadata information from file providing similar output to what Explorer shows when viewing file
    .PARAMETER File
    FileName or FileObject
    .EXAMPLE
    Get-ChildItem -Path $Env:USERPROFILE\Desktop -Force | Get-FileMetaData | Out-HtmlView -ScrollX -Filtering -AllProperties
    .EXAMPLE
    Get-ChildItem -Path $Env:USERPROFILE\Desktop -Force | Where-Object { $_.Attributes -like '*Hidden*' } | Get-FileMetaData | Out-HtmlView -ScrollX -Filtering -AllProperties
    .NOTES
        Written by Lukas Woehrl (woehrl01)
        Modified by Skyler Werner
        Version: 2.0.0
    #>
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, ValueFromPipeline)]
        [Object]$File,
        [Switch]$Signature
    )

process {

    foreach ($f in $File) {

        $metaDataObject = [ordered] @{}
        if ($f -is [string]) {
            $fileInformation = Get-ItemProperty -Path $f
        }
        elseif ($f -is [System.IO.DirectoryInfo]) {
            # Write-Warning "Get-FileMetaData - Directories are not supported. Skipping $f."
            continue
        }
        elseif ($f -is [System.IO.FileInfo]) {
            $fileInformation = $f
        }
        else {
            Write-Warning "Get-FileMetaData - Only files are supported. Skipping $f."
            continue
        }

        $shellApplication = New-Object -ComObject Shell.Application
        $shellFolder = $shellApplication.Namespace($fileInformation.Directory.FullName)
        $shellFile = $shellFolder.ParseName($fileInformation.Name)
        $metaDataProperties = [ordered] @{}

        0..400 | ForEach-Object -Process {
            $dataValue = $shellFolder.GetDetailsOf($null, $_)
            $propertyValue = (Get-Culture).TextInfo.ToTitleCase($dataValue.Trim()).Replace(' ', '')

            if ($propertyValue -ne '') {
                $metaDataProperties["$_"] = $propertyValue
            }
        }

        foreach ($key in $metaDataProperties.Keys) {
            $property = $metaDataProperties[$key]
            $value = $shellFolder.GetDetailsOf($shellFile, [int] $key)
            if ($property -in 'Attributes', 'Folder', 'Type', 'SpaceFree', 'TotalSize', 'SpaceUsed') {
                continue
            }
            if (($null -ne $value) -and ($value -ne '')) {
                $metaDataObject["$property"] = $value
            }
        }

        if ($fileInformation.VersionInfo) {
            $splitInfo = ([string] $fileInformation.VersionInfo).Split([char]13)
            foreach ($item in $splitInfo) {
                $property = $item.Split(":").Trim()

                if ($property[0] -and $property[1] -ne '') {
                    $metaDataObject["$($property[0])"] = $property[1]
                }
            }
        }

        $metaDataObject["Attributes"] = $fileInformation.Attributes
        $metaDataObject['IsReadOnly'] = $fileInformation.IsReadOnly
        $metaDataObject['IsHidden'] = $fileInformation.Attributes -like '*Hidden*'
        $metaDataObject['IsSystem'] = $fileInformation.Attributes -like '*System*'

        if ($Signature) {
            $digitalSignature = Get-AuthenticodeSignature -FilePath $fileInformation.Fullname
            $metaDataObject['SignatureCertificateSubject'] = $digitalSignature.SignerCertificate.Subject
            $metaDataObject['SignatureCertificateIssuer'] = $digitalSignature.SignerCertificate.Issuer
            $metaDataObject['SignatureCertificateSerialNumber'] = $digitalSignature.SignerCertificate.SerialNumber
            $metaDataObject['SignatureCertificateNotBefore'] = $digitalSignature.SignerCertificate.NotBefore
            $metaDataObject['SignatureCertificateNotAfter'] = $digitalSignature.SignerCertificate.NotAfter
            $metaDataObject['SignatureCertificateThumbprint'] = $digitalSignature.SignerCertificate.Thumbprint
            $metaDataObject['SignatureStatus'] = $digitalSignature.Status
            $metaDataObject['IsOSBinary'] = $digitalSignature.IsOSBinary
        }

        [PSCustomObject]$metaDataObject
    }

} # End process

} # End function

Export-ModuleMember Get-FileMetaData
