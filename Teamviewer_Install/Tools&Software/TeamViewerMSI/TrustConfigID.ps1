param
(
    [Parameter(Mandatory=$False)]
    [string]$Add,
    [Parameter(Mandatory=$False)]
    [string]$Path,
    [Parameter(Mandatory=$False)]
    [string]$Name,
    [Parameter(Mandatory=$False)]
    [string]$Remove,
    [Parameter(Mandatory=$False)]
    [switch]$List,
    [Parameter(Mandatory=$False)]
    [switch]$Merge
)

If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{   
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Break
}

function Load-Module($name)
{
    if (-not(Get-Module -Name $name))
    {
        if (Get-Module -ListAvailable | Where-Object { $_.name -eq $name })
        {
            Import-Module $name  

            return $true
        }
        else
        {   
            return $false
        }
    }
    else
    {
        return $true
    }
}

$moduleName = "ActiveDirectory"

if (-not(Load-Module $moduleName))
{
    Write-Host "Failed to load $moduleName"
    Write-Host "Please download the Active Directory Module from microsoft.com"
    Read-Host
    Break
}


[string]$TV_GUID = "B15CB251-377F-46FB-81E9-4B6F12D6A15F"
[int]$StringCutStartIndex = 6


###########################
#Functions
###########################

################
#PRIVATE
################

#The Main function, checks how many TeamViewer ADObjects exists and process the input
function Main
{
    #If more than one TV SCP were found -> error
    if((Count-TV_SCPs) -gt 1 -and !$Merge)
    {
        Write-Error "Too many ServerConnectionPoints with TeamViewer keywords are existing!"
        List-TV_SCPs
    }
    else
    {
        $HasSomethingDone = $False
        if($Add)
        {
            Add-ConfigurationID -ConfigurationID $Add -Name $Name -Path $Path
            $HasSomethingDone = $True
        }
        if($Remove)
        {
            Remove -ConfigurationID $Remove
            $HasSomethingDone = $True
        }
        if($List)
        {
            List-ConfigurationIDs
            $HasSomethingDone = $True
        }
        if($Merge)
        {
            Merge -Name $Name -Path $Path
            $HasSomethingDone = $True
        }
        if(!$HasSomethingDone)
        {
            do
            {
                $ConfigurationID = Read-Host "Please enter a ConfigurationID to add"
            }while(!$ConfigurationID)
            Add-ConfigurationID -ConfigurationID $ConfigurationID
            Read-Host "Enter any key to continue..."
        }
    }
    
}

#Returns the TV SCP(s)
function Get-TV_SCPs
{
    return Get-ADObject -Filter {objectClass -like "serviceConnectionPoint" -and keywords -like 'TeamViewer' -and keywords -like $TV_GUID}
}

#Lists all ServiceConnectionPoints with TeamViewer keyword
function List-TV_SCPs
{
    if($TeamViewerTV_SCPs = Get-TV_SCPs)
    {
        foreach($CurTV_SCP in $TeamViewerTV_SCPs)
        {
            $CurTV_SCP
        }
    }
    return 
}

#Counts all ServiceConnectionPoints with TeamViewer keyword
function Count-TV_SCPs
{
    $Counter = 0
    if($TeamViewerTV_SCPs = Get-TV_SCPs)
    {
        foreach($CurTV_SCP in $TeamViewerTV_SCPs)
        {
            $Counter++
        }
    }
    return $Counter
}

#Counts all ConfigurationIDs saved in the TeamViewer ServiceConnectionPoint
function Count-ConfigurationIDs
{
    $Counter = 0
    if($TeamViewerTV_SCP = Get-TV_SCPs)
    {
        [string]$ServiceBindingInformation = (Get-ADObject $TeamViewerTV_SCP -Properties serviceBindingInformation | select @{name="SBI";expression={$_.serviceBindingInformation -join “;”}})
        if($ServiceBindingInformation)
        {
            $ServiceBindingInformation = $ServiceBindingInformation.Substring($StringCutStartIndex, $ServiceBindingInformation.Length - $StringCutStartIndex - 1)
            $SplittedSBI = $ServiceBindingInformation.Split(';')
            foreach($CurValue in $SplittedSBI)
            {
                if($CurValue.Length -gt 1)
                {
                    $Counter++
                }
            }
        }
    }
    return $Counter
}

#Creates the ServiceConnenctionPoint with the TeamViewer keywords
function Create-TV_SCP
{
    param
    (
        [string]$Name,
        [string]$SCP_Path,
        [switch]$Automatic
    )
    if(!$Name)
    {
        $Name = Read-Host "Name"
    }
    if(!$SCP_Path)
    {
        $SCP_Path = Read-Host "Path"
    }
    $ContainerPath = $SCP_Path
    $ContainerFullName = $SCP_Path
    if(!$SCP_Path)
    {
        $SCP_Path = Get-ADDomain
        $ContainerPath =  "CN=System,"+$SCP_Path
        $ContainerFullName = "CN=TeamViewer,"+$ContainerPath
        $SCP_Path = $ContainerFullName
    }
    if(!$Name)
    {
        $Name = "TeamViewer"
    }
    if(!(Get-ADObject -Filter { distinguishedName -like $ContainerFullName -and objectClass -like "container"}))
    {
        $ContainerName = "TeamViewer"
        New-ADObject -Name $ContainerName -Path $ContainerPath -Type container
        "Container with name '"+$ContainerName+"' was created because the TeamViewer-ServiceConnectionPoint must be in a container."
    }
    #Create the TV SCP
    New-ADObject -Name $Name -Path $SCP_Path -Type serviceConnectionPoint -OtherAttributes @{'keywords'=$TV_GUID,"TeamViewer"}
    "TeamViewer-ServiceConnectionPoint with name '"+$Name+"' was created"
    return $SCP_Path
}

################
#ACTIONS
################

#Lists all ConfigurationIDs saved in the TeamViewer ServiceConnectionPoint
function List-ConfigurationIDs
{
    if($TeamViewerTV_SCP = Get-TV_SCPs)
    {
        [string]$ServiceBindingInformation = (Get-ADObject $TeamViewerTV_SCP -Properties serviceBindingInformation | select @{name="SBI";expression={$_.serviceBindingInformation -join “;”}})
        if($ServiceBindingInformation)
        {
            $ServiceBindingInformation = $ServiceBindingInformation.Substring($StringCutStartIndex, $ServiceBindingInformation.Length - $StringCutStartIndex - 1)
            $SplittedSBI = $ServiceBindingInformation.Split(';')
            foreach($CurValue in $SplittedSBI)
            {
                $CurValue
            }
        }
    }
    return
}

#Merge all ServiceConnectionPoints with TeamViewer keyword to one and combine their saved ConfigurationIDs
function Merge
{
    if(!$Path -and $Name)
    {
        $Path = Get-ADDomain
        $Path = "CN="+$Name+"_Container,"+$Path
    }
    if($TeamViewerTV_SCPs = Get-TV_SCPs)
    {
        #Declare $FirstTV_SCP for further using
        $FirstTV_SCP
        [bool]$FirstChosen = $false

        #Check if TV_SCP with given name already exists
        $CurTV_SCP = Get-TV_SCPs
        $CurTV_SCP = $CurTV_SCP | Where-Object {$_.Name -like $Name}
        if($CurTV_SCP)
        {
            $FirstTV_SCP = $CurTV_SCP
            $FirstChosen = $true
        }
        #If not but name and path are given, create a new
        if($Name -and $Path -and !$FirstChosen)
        {
            $Path = Create -Name $Name -Path $Path
            $FirstTV_SCP = "CN="+$Name+","+$Path
            $FirstChosen = $true
        }
        
        foreach($CurTV_SCP in $TeamViewerTV_SCPs)
        {
            #If no TV_SCP was chosen, use the first found
            if(!$FirstChosen)
            {
                $FirstTV_SCP = $CurTV_SCP
                $FirstChosen = $true
            }
            else
            {
                if(!($CurTV_SCP -like $FirstTV_SCP))
                {
                    [string]$ServiceBindingInformation = (Get-ADObject $CurTV_SCP -Properties serviceBindingInformation | select @{name="SBI";expression={$_.serviceBindingInformation -join “;”}})
                    if($ServiceBindingInformation)
                    {
                        $ServiceBindingInformation = $ServiceBindingInformation.Substring($StringCutStartIndex, $ServiceBindingInformation.Length - $StringCutStartIndex - 1)
                        $SplittedSBI = $ServiceBindingInformation.Split(';')
                        foreach($CurValue in $SplittedSBI)
                        {
                            Set-ADObject -Identity $FirstTV_SCP -Add @{'serviceBindingInformation' = $CurValue}
                        }
                    }
                    Remove-ADObject -Identity $CurTV_SCP
                }
            }

        }
    }
}



#Creates a TeamViewer ServiceConnectionPoint if it doesn't exits and adds a ConfigurationID to it
function Add-ConfigurationID
{
    param
    (
        [Parameter(Mandatory=$True)]
        [string]$ConfigurationID,
        [Parameter(Mandatory=$False)]
        [string]$Path,
        [Parameter(Mandatory=$False)]
        [string]$Name            
    )
    if((Count-TV_SCPs) -lt 1)
    {
        Create-TV_SCP -Name $Name -SCP_Path $Path
    }
    $TV_SCP = Get-TV_SCPs
    Set-ADObject -Identity $TV_SCP -Add @{'serviceBindingInformation' = $ConfigurationID}
    Write-Host "Successfully added Configuration ID: ",$ConfigurationID
}

#Removes a ConfigurationID from the TeamViewer ServiceConnectionPoint and destroy it if it was the last ConfigurationID
function Remove
{
    param
    (
        [Parameter(Mandatory=$False)]
        [string]$ConfigurationID       
    )
    
    if((Count-TV_SCPs) -eq 1)
    {
        [string]$UpperCasedOptions = $ConfigurationID.ToUpper()
        $TV_SCP = Get-TV_SCPs
        if($UpperCasedOptions -eq "ALL")
        {
            Remove-ADObject -Identity $TV_SCP
        }
        else
        {
            Set-ADObject -Identity $TV_SCP -Remove @{'serviceBindingInformation' = $ConfigurationID}
            Write-Host "Successfully removed Configuration ID: ",$ConfigurationID
            if((Count-ConfigurationIDs) -lt 1)
            {
                Remove-ADObject -Identity $TV_SCP
            }
        }
    }
    else
    {
        Write-Host "No TeamViewer ServiceConnectionPoint was found!"
    }
}


#Call the Main function
Main


     




# SIG # Begin signature block
# MIIZIQYJKoZIhvcNAQcCoIIZEjCCGQ4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUuKp5d7w5FoXNBTXbJzY7qwtO
# eZygghPiMIID7jCCA1egAwIBAgIQfpPr+3zGTlnqS5p31Ab8OzANBgkqhkiG9w0B
# AQUFADCBizELMAkGA1UEBhMCWkExFTATBgNVBAgTDFdlc3Rlcm4gQ2FwZTEUMBIG
# A1UEBxMLRHVyYmFudmlsbGUxDzANBgNVBAoTBlRoYXd0ZTEdMBsGA1UECxMUVGhh
# d3RlIENlcnRpZmljYXRpb24xHzAdBgNVBAMTFlRoYXd0ZSBUaW1lc3RhbXBpbmcg
# Q0EwHhcNMTIxMjIxMDAwMDAwWhcNMjAxMjMwMjM1OTU5WjBeMQswCQYDVQQGEwJV
# UzEdMBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xMDAuBgNVBAMTJ1N5bWFu
# dGVjIFRpbWUgU3RhbXBpbmcgU2VydmljZXMgQ0EgLSBHMjCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBALGss0lUS5ccEgrYJXmRIlcqb9y4JsRDc2vCvy5Q
# WvsUwnaOQwElQ7Sh4kX06Ld7w3TMIte0lAAC903tv7S3RCRrzV9FO9FEzkMScxeC
# i2m0K8uZHqxyGyZNcR+xMd37UWECU6aq9UksBXhFpS+JzueZ5/6M4lc/PcaS3Er4
# ezPkeQr78HWIQZz/xQNRmarXbJ+TaYdlKYOFwmAUxMjJOxTawIHwHw103pIiq8r3
# +3R8J+b3Sht/p8OeLa6K6qbmqicWfWH3mHERvOJQoUvlXfrlDqcsn6plINPYlujI
# fKVOSET/GeJEB5IL12iEgF1qeGRFzWBGflTBE3zFefHJwXECAwEAAaOB+jCB9zAd
# BgNVHQ4EFgQUX5r1blzMzHSa1N197z/b7EyALt0wMgYIKwYBBQUHAQEEJjAkMCIG
# CCsGAQUFBzABhhZodHRwOi8vb2NzcC50aGF3dGUuY29tMBIGA1UdEwEB/wQIMAYB
# Af8CAQAwPwYDVR0fBDgwNjA0oDKgMIYuaHR0cDovL2NybC50aGF3dGUuY29tL1Ro
# YXd0ZVRpbWVzdGFtcGluZ0NBLmNybDATBgNVHSUEDDAKBggrBgEFBQcDCDAOBgNV
# HQ8BAf8EBAMCAQYwKAYDVR0RBCEwH6QdMBsxGTAXBgNVBAMTEFRpbWVTdGFtcC0y
# MDQ4LTEwDQYJKoZIhvcNAQEFBQADgYEAAwmbj3nvf1kwqu9otfrjCR27T4IGXTdf
# plKfFo3qHJIJRG71betYfDDo+WmNI3MLEm9Hqa45EfgqsZuwGsOO61mWAK3ODE2y
# 0DGmCFwqevzieh1XTKhlGOl5QGIllm7HxzdqgyEIjkHq3dlXPx13SYcqFgZepjhq
# IhKjURmDfrYwggSjMIIDi6ADAgECAhAOz/Q4yP6/NW4E2GqYGxpQMA0GCSqGSIb3
# DQEBBQUAMF4xCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBDb3Jwb3Jh
# dGlvbjEwMC4GA1UEAxMnU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBD
# QSAtIEcyMB4XDTEyMTAxODAwMDAwMFoXDTIwMTIyOTIzNTk1OVowYjELMAkGA1UE
# BhMCVVMxHTAbBgNVBAoTFFN5bWFudGVjIENvcnBvcmF0aW9uMTQwMgYDVQQDEytT
# eW1hbnRlYyBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIFNpZ25lciAtIEc0MIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAomMLOUS4uyOnREm7Dv+h8GEKU5Ow
# mNutLA9KxW7/hjxTVQ8VzgQ/K/2plpbZvmF5C1vJTIZ25eBDSyKV7sIrQ8Gf2Gi0
# jkBP7oU4uRHFI/JkWPAVMm9OV6GuiKQC1yoezUvh3WPVF4kyW7BemVqonShQDhfu
# ltthO0VRHc8SVguSR/yrrvZmPUescHLnkudfzRC5xINklBm9JYDh6NIipdC6Anqh
# d5NbZcPuF3S8QYYq3AhMjJKMkS2ed0QfaNaodHfbDlsyi1aLM73ZY8hJnTrFxeoz
# C9Lxoxv0i77Zs1eLO94Ep3oisiSuLsdwxb5OgyYI+wu9qU+ZCOEQKHKqzQIDAQAB
# o4IBVzCCAVMwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAO
# BgNVHQ8BAf8EBAMCB4AwcwYIKwYBBQUHAQEEZzBlMCoGCCsGAQUFBzABhh5odHRw
# Oi8vdHMtb2NzcC53cy5zeW1hbnRlYy5jb20wNwYIKwYBBQUHMAKGK2h0dHA6Ly90
# cy1haWEud3Muc3ltYW50ZWMuY29tL3Rzcy1jYS1nMi5jZXIwPAYDVR0fBDUwMzAx
# oC+gLYYraHR0cDovL3RzLWNybC53cy5zeW1hbnRlYy5jb20vdHNzLWNhLWcyLmNy
# bDAoBgNVHREEITAfpB0wGzEZMBcGA1UEAxMQVGltZVN0YW1wLTIwNDgtMjAdBgNV
# HQ4EFgQURsZpow5KFB7VTNpSYxc/Xja8DeYwHwYDVR0jBBgwFoAUX5r1blzMzHSa
# 1N197z/b7EyALt0wDQYJKoZIhvcNAQEFBQADggEBAHg7tJEqAEzwj2IwN3ijhCcH
# bxiy3iXcoNSUA6qGTiWfmkADHN3O43nLIWgG2rYytG2/9CwmYzPkSWRtDebDZw73
# BaQ1bHyJFsbpst+y6d0gxnEPzZV03LZc3r03H0N45ni1zSgEIKOq8UvEiCmRDoDR
# EfzdXHZuT14ORUZBbg2w6jiasTraCXEQ/Bx5tIB7rGn0/Zy2DBYr8X9bCT2bW+IW
# yhOBbQAuOA2oKY8s4bL0WqkBrxWcLC9JG9siu8P+eJRRw4axgohd8D20UaF5Mysu
# e7ncIAkTcetqGVvP6KUwVyyJST+5z3/Jvz4iaGNTmr1pdKzFHTx/kuDDvBzYBHUw
# ggU3MIIEH6ADAgECAhBWcpMAx4MGxCZ8pEoQrc0DMA0GCSqGSIb3DQEBBQUAMIG0
# MQswCQYDVQQGEwJVUzEXMBUGA1UEChMOVmVyaVNpZ24sIEluYy4xHzAdBgNVBAsT
# FlZlcmlTaWduIFRydXN0IE5ldHdvcmsxOzA5BgNVBAsTMlRlcm1zIG9mIHVzZSBh
# dCBodHRwczovL3d3dy52ZXJpc2lnbi5jb20vcnBhIChjKTEwMS4wLAYDVQQDEyVW
# ZXJpU2lnbiBDbGFzcyAzIENvZGUgU2lnbmluZyAyMDEwIENBMB4XDTE0MDczMDAw
# MDAwMFoXDTE3MDkwNTIzNTk1OVowaTELMAkGA1UEBhMCREUxGzAZBgNVBAgTEkJh
# ZGVuIFd1ZXJ0dGVtYmVyZzETMBEGA1UEBxMKR29lcHBpbmdlbjETMBEGA1UEChQK
# VGVhbVZpZXdlcjETMBEGA1UEAxQKVGVhbVZpZXdlcjCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAKUOMmKuX5le+2bjU4BdCUHzzN4Ysl7ZL7XqQPilELxt
# phaBdLE+laBpNcmKYJP2XZoPxRS28g+bCYZKSC62h/kRQNTF/OhZZj/kn2ppvhRK
# QNbbp6q3ofT8bqhNRxNUIpUnKTtTOlKNfJ2rzoQXvJ0ZcCWOqg8To8T8dUXY4vpp
# aCx+OdZlfjvv4pUN5H6Mjk/Byot3bfgJ0kigmplavJ01FbJ0AvxKc+VB0s2DC/bb
# N57W7VwD9bHCbSjsmzavVzUefes8qwnflEbWRgB075DKJR4TfcUAfxwkdd8Grhwx
# 4ZgD7yZashkTC6UUIKf+CRYVB1XjUTAj6+zJneZXaZcCAwEAAaOCAY0wggGJMAkG
# A1UdEwQCMAAwDgYDVR0PAQH/BAQDAgeAMCsGA1UdHwQkMCIwIKAeoByGGmh0dHA6
# Ly9zZi5zeW1jYi5jb20vc2YuY3JsMGYGA1UdIARfMF0wWwYLYIZIAYb4RQEHFwMw
# TDAjBggrBgEFBQcCARYXaHR0cHM6Ly9kLnN5bWNiLmNvbS9jcHMwJQYIKwYBBQUH
# AgIwGRYXaHR0cHM6Ly9kLnN5bWNiLmNvbS9ycGEwEwYDVR0lBAwwCgYIKwYBBQUH
# AwMwVwYIKwYBBQUHAQEESzBJMB8GCCsGAQUFBzABhhNodHRwOi8vc2Yuc3ltY2Qu
# Y29tMCYGCCsGAQUFBzAChhpodHRwOi8vc2Yuc3ltY2IuY29tL3NmLmNydDAfBgNV
# HSMEGDAWgBTPmanqeyb0S8mOj9fwBSbv49KnnTAdBgNVHQ4EFgQUXd442PTuZGnT
# Fk9y8ZfhGcPqPPowEQYJYIZIAYb4QgEBBAQDAgQQMBYGCisGAQQBgjcCARsECDAG
# AQEAAQH/MA0GCSqGSIb3DQEBBQUAA4IBAQC0HqHiBidtAoyGJXKn4+cvR0m/zPnE
# wsWteIKads8fUi2CCi+BpfCzFiscz4MtIH4VpygddqPfbPxGZxLp9IzbpBRiBakC
# 1cHI8H47wkKtBk24j3JacA0BtYJGoBfL/cv9bQc+mPC0yOtIlnOCFYLQEcI1FhZ7
# jIAaPtP/FcL4X+Vjq1/TcuSA7Cii+xfKcekIYRXz9nfmTEL30cM4l05aC9+cvW6h
# 7TxTrG+HDZOZSLUVRaqBrb8nihZSgKeQZqW+/TYTxtQm0iMWbZ68nvZMdJxQeYQy
# Ht+LmV2ChObvnNUnFyQVS0fmfG83RgL8PjKdgh901z4d+MDZwa89XqoYMIIGCjCC
# BPKgAwIBAgIQUgDlqiVW/BqG7ZbJ1EszxzANBgkqhkiG9w0BAQUFADCByjELMAkG
# A1UEBhMCVVMxFzAVBgNVBAoTDlZlcmlTaWduLCBJbmMuMR8wHQYDVQQLExZWZXJp
# U2lnbiBUcnVzdCBOZXR3b3JrMTowOAYDVQQLEzEoYykgMjAwNiBWZXJpU2lnbiwg
# SW5jLiAtIEZvciBhdXRob3JpemVkIHVzZSBvbmx5MUUwQwYDVQQDEzxWZXJpU2ln
# biBDbGFzcyAzIFB1YmxpYyBQcmltYXJ5IENlcnRpZmljYXRpb24gQXV0aG9yaXR5
# IC0gRzUwHhcNMTAwMjA4MDAwMDAwWhcNMjAwMjA3MjM1OTU5WjCBtDELMAkGA1UE
# BhMCVVMxFzAVBgNVBAoTDlZlcmlTaWduLCBJbmMuMR8wHQYDVQQLExZWZXJpU2ln
# biBUcnVzdCBOZXR3b3JrMTswOQYDVQQLEzJUZXJtcyBvZiB1c2UgYXQgaHR0cHM6
# Ly93d3cudmVyaXNpZ24uY29tL3JwYSAoYykxMDEuMCwGA1UEAxMlVmVyaVNpZ24g
# Q2xhc3MgMyBDb2RlIFNpZ25pbmcgMjAxMCBDQTCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBAPUjS16l14q7MunUV/fv5Mcmfq0ZmP6onX2U9jZrENd1gTB/
# BGh/yyt1Hs0dCIzfaZSnN6Oce4DgmeHuN01fzjsU7obU0PUnNbwlCzinjGOdF6MI
# pauw+81qYoJM1SHaG9nx44Q7iipPhVuQAU/Jp3YQfycDfL6ufn3B3fkFvBtInGnn
# wKQ8PEEAPt+W5cXklHHWVQHHACZKQDy1oSapDKdtgI6QJXvPvz8c6y+W+uWHd8a1
# VrJ6O1QwUxvfYjT/HtH0WpMoheVMF05+W/2kk5l/383vpHXv7xX2R+f4GXLYLjQa
# prSnTH69u08MPVfxMNamNo7WgHbXGS6lzX40LYkCAwEAAaOCAf4wggH6MBIGA1Ud
# EwEB/wQIMAYBAf8CAQAwcAYDVR0gBGkwZzBlBgtghkgBhvhFAQcXAzBWMCgGCCsG
# AQUFBwIBFhxodHRwczovL3d3dy52ZXJpc2lnbi5jb20vY3BzMCoGCCsGAQUFBwIC
# MB4aHGh0dHBzOi8vd3d3LnZlcmlzaWduLmNvbS9ycGEwDgYDVR0PAQH/BAQDAgEG
# MG0GCCsGAQUFBwEMBGEwX6FdoFswWTBXMFUWCWltYWdlL2dpZjAhMB8wBwYFKw4D
# AhoEFI/l0xqGrI2Oa8PPgGrUSBgsexkuMCUWI2h0dHA6Ly9sb2dvLnZlcmlzaWdu
# LmNvbS92c2xvZ28uZ2lmMDQGA1UdHwQtMCswKaAnoCWGI2h0dHA6Ly9jcmwudmVy
# aXNpZ24uY29tL3BjYTMtZzUuY3JsMDQGCCsGAQUFBwEBBCgwJjAkBggrBgEFBQcw
# AYYYaHR0cDovL29jc3AudmVyaXNpZ24uY29tMB0GA1UdJQQWMBQGCCsGAQUFBwMC
# BggrBgEFBQcDAzAoBgNVHREEITAfpB0wGzEZMBcGA1UEAxMQVmVyaVNpZ25NUEtJ
# LTItODAdBgNVHQ4EFgQUz5mp6nsm9EvJjo/X8AUm7+PSp50wHwYDVR0jBBgwFoAU
# f9Nlp8Ld7LvwMAnzQzn6Aq8zMTMwDQYJKoZIhvcNAQEFBQADggEBAFYi5jSkxGHL
# SLkBrVaoZA/ZjJHEu8wM5a16oCJ/30c4Si1s0X9xGnzscKmx8E/kDwxT+hVe/nSY
# SSSFgSYckRRHsExjjLuhNNTGRegNhSZzA9CpjGRt3HGS5kUFYBVZUTn8WBRr/tSk
# 7XlrCAxBcuc3IgYJviPpP0SaHulhncyxkFz8PdKNrEI9ZTbUtD1AKI+bEM8jJsxL
# IMuQH12MTDTKPNjlN9ZvpSC9NOsm2a4N58Wa96G0IZEzb4boWLslfHQOWP51G2M/
# zjF8m48blp7FU3aEW5ytkfqs7ZO6XcghU8KCU2OvEg1QhxEbPVRSloosnD2SGgia
# BS7Hk6VIkdMxggSpMIIEpQIBATCByTCBtDELMAkGA1UEBhMCVVMxFzAVBgNVBAoT
# DlZlcmlTaWduLCBJbmMuMR8wHQYDVQQLExZWZXJpU2lnbiBUcnVzdCBOZXR3b3Jr
# MTswOQYDVQQLEzJUZXJtcyBvZiB1c2UgYXQgaHR0cHM6Ly93d3cudmVyaXNpZ24u
# Y29tL3JwYSAoYykxMDEuMCwGA1UEAxMlVmVyaVNpZ24gQ2xhc3MgMyBDb2RlIFNp
# Z25pbmcgMjAxMCBDQQIQVnKTAMeDBsQmfKRKEK3NAzAJBgUrDgMCGgUAoIGmMBkG
# CSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEE
# AYI3AgEVMCMGCSqGSIb3DQEJBDEWBBQjb8MLqHkAen4AedrICyoeWbv5GTBGBgor
# BgEEAYI3AgEMMTgwNqAWgBQAVABlAGEAbQBWAGkAZQB3AGUAcqEcgBpodHRwOi8v
# d3d3LnRlYW12aWV3ZXIuY29tIDANBgkqhkiG9w0BAQEFAASCAQBYWoNgEfTDnsOd
# AioDJfQm120NxPMPCtaNAiKr9DBYY6WbnuA68xEtRVZ+sk9oB6MFMzHheAUdCl8i
# tMjNr/Bosqkm6t+ZbpB9KHgqdSOOJpUYt9sP4VJj9FbTh/54/9UlBl0//Uyi6I2G
# PYj2UaykZxIdrSrdT97XVEk0CjLkGxR9ugzwrGQU5jrT+5GuT4VgPvmCmzUctA/w
# eVod2gDiHGkBplzrH6r+JcQuo/U5qlC3YhfLbV5bLOPa811lIhGROt9ar26NF4up
# m66RixuWD9lBnrA9FRshmt9cCif1y47CDyXOKWBlgfvAV4dpvosce+JrMDjm8MOO
# SDbtmvVZoYICCzCCAgcGCSqGSIb3DQEJBjGCAfgwggH0AgEBMHIwXjELMAkGA1UE
# BhMCVVMxHTAbBgNVBAoTFFN5bWFudGVjIENvcnBvcmF0aW9uMTAwLgYDVQQDEydT
# eW1hbnRlYyBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIENBIC0gRzICEA7P9DjI/r81
# bgTYapgbGlAwCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEw
# HAYJKoZIhvcNAQkFMQ8XDTE1MDkxMTE2MzgzN1owIwYJKoZIhvcNAQkEMRYEFG5h
# WhIdOAy/jxfgnoLR2VG4L/NyMA0GCSqGSIb3DQEBAQUABIIBABNcBUzhjCNzpGzr
# C5Guq6TK4RhrDCx5XmMaNp0ydEMqBhqmkLMtdoGXVp2CEc1YwSCfTVxbHffokD7L
# a4BITLYDYdW6MpWYmohzWx99RKGHi3DS3MZg8PnTOYh1MC0298eXcWOfSi2BI5b4
# tgUaQhp2yOpc8zIY3JxA+yLy/FLgbVHcwrAHZweVSnIc6phsFI51QUkPvN9fhsUF
# iDwLKh4vl2drZL9/tRoOrgPMaehPPtboXQx13//aS+Ua+/ahVq4ZdUtbrHeD9mfo
# xyv9zlAexFWDwvxhaQc+UWC6hfs/Q7fVBZ7OWC2ATVvmHyh9d+hZvaQwlPp8J7rX
# ZBTHEi8=
# SIG # End signature block
