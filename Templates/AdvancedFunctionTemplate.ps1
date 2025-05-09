function New-AdvancedFunction{
    <#
    .Synopsis
        Short Description
    .DESCRIPTION
        Long Description
    .EXAMPLE
        PS> New-AdvancedFunctionTemplate -Param1 MYPARAM
    .EXAMPLE
        Another example of how to use this cmdlet
    .PARAMETER Param1
        This param does this thing.
    .PARAMETER
    .PARAMETER
    .PARAMETER
    #>

    #Script Name:
    #Created by Will Hughes
    #Date: 
    #Patch Notes:


    [CmdletBinding()]
    param (
        [string]$Param1
    )
    process {
        try {
            # Your code here
        }
        catch {
            Write-Error "$($_.Exception.Message) - Line Number: $($_.InvocationInfo.ScriptLineNumber)"
        }
    }
}