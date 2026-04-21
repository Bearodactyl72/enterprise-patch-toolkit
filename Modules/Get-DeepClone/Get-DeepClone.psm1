# DOTS formatting comment

<#
    .SYNOPSIS
        Returns a new hash table completely unlinked to an original hash table.
    .DESCRIPTION
        Returns a new hash table completely unlinked to an original hash table input.
        Written by Skyler Werner
        Date: 2021/05/05
        Modified: 2024/03/12
        Version 1.0.1
#>

function Get-DeepClone {
    [cmdletbinding()]
    param(
        $InputObject
    )

    process {
        if($InputObject -is [Hashtable]) {
            $clone = @{}
            foreach ($key in $InputObject.Keys) {
                # This is a recursive function because it references itself
                $clone[$key] = Get-DeepClone $InputObject[$key]
            }
            return $clone
        }
        else {
            return $InputObject
        }
    }
}

Export-ModuleMember Get-DeepClone
