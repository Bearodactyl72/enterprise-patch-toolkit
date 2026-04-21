# DOTS formatting comment

<#
    .SYNOPSIS
        Adds a delimiter to a string that contains numbers separated by a space.
    .DESCRIPTION
        Adds a delimiter to a string that contains numbers separated by a space.
        Written by Skyler Werner
        Date: 2021/05/05
        Modified: 2024/03/15
        Version 1.0.1
#>

function Add-Delimiter {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        $InputObject,

        # Properties of the input object that require a delimiter
        [Parameter(Mandatory = $true, Position = 1)]
        [String[]]$Property,

        # Default delimiter is a comma
        [Parameter(Position = 2)]
        [string]$Delimiter = ","
    )

    foreach ($inputItem in $InputObject) {

        $inputClone = $inputItem.PsObject.Copy()

        foreach ($prop in $Property) {

            # Skip properties that don't exist on this object (e.g. incomplete results)
            if ($null -eq $inputClone.PSObject.Properties[$prop]) { continue }

            $inputClone.$prop = [string]$inputClone.$prop

            if (($inputClone.$prop -match '\d') -and ($inputClone.$prop -match " ")) {
                $inputClone.$prop = $inputClone.$prop.Replace(" ","$Delimiter ")
            }
        }
        $inputClone
    }
}

Export-ModuleMember Add-Delimiter
