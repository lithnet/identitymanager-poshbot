# Slack text width with the formatting we use maxes out ~80 characters...
$Width = 80
$CommandsToExport = @()

function Get-MVStats {
    <#
    .SYNOPSIS
        Get MIM Sync Engine Stats
    .EXAMPLE
        !MVStats
    .EXAMPLE
        !Stats
    .DESCRIPTION !Stats
        MIM Sync Engine MA Stats
    .LINK
        https://blog.darrenjrobinson.com
    #>
    [cmdletbinding()]
    [PoshBot.BotCommand(
        CommandName = 'MVStats',
        Aliases = ('Stats', 'Get-MVStats'),
        Permissions = 'read'
    )]
    param(
        # !Stats 
        [PoshBot.FromConfig('MIMSyncCreds')]
        [parameter(mandatory)]
        [PSCredential]$MIMSyncCreds,

        [PoshBot.FromConfig('MIMSyncAddress')]
        [parameter(mandatory)]
        [string]$MIMSyncAddress
    )
    
    # Remote PowerShell Scriptblock
    $scriptblock = {
        # Import LithnetMIISAutomation for MIM Sync Server Config Exports
        # It must be installed on the MIM Sync Server
        Import-Module lithnetmiisautomation; 
        # Query MAs
        $MAs = Get-ManagementAgent
        $intTotalConnectors = 0
        $intTotalObjects = 0
        $MAStats = @()

        if ($MAs) {
            foreach ($ma in $MAs) {
                $objects = $ma.Statistics.Total
                $connectors = $ma.Statistics.TotalConnectors
                $lastrun = Get-LastRunDetails -MA $ma | Select-Object StartTime, EndTime
                $lastrunEnd = $lastrun.EndTime
                $intTotalConnectors += $connectors
                $intTotalObjects += $objects

                # Output Stats for the report
                $maAttr = New-Object -TypeName PSObject
                $maAttr | Add-Member -Type NoteProperty -Name "Management Agent" -Value $ma.Name
                $maAttr | Add-Member -Type NoteProperty -Name "Total Objects" -Value $objects
                $maAttr | Add-Member -Type NoteProperty -Name "Connectors" -Value $connectors 
                $maAttr | Add-Member -Type NoteProperty -Name "Last Run" -Value $lastrunEnd

                $MAStats += $maAttr      
            }

            # Totals
            $maAttr = New-Object -TypeName PSObject
            $maAttr | Add-Member -Type NoteProperty -Name "Management Agent" -Value "Total"
            $maAttr | Add-Member -Type NoteProperty -Name "Total Objects" -Value $intTotalObjects 
            $maAttr | Add-Member -Type NoteProperty -Name "Connectors" -Value $intTotalConnectors 
            # Sort by MA Name
            $MAStats = $MAStats | Sort-Object -Property "Management Agent"
            # Then add the Total to the end 
            $MAStats += $maAttr 
            # Output
            $MAStats 
        }
        else {
            Write-Output "Connection to MIM Sync Server Failed"
        }
    }

    # Connect to MIM Sync Server using Remote PowerShell to run LithnetMIISAutomation cmdlets   
    $result = invoke-command -ScriptBlock $scriptblock -ComputerName $MIMSyncAddress -Authentication Kerberos -Credential $MIMSyncCreds 
    $o = $result | Select-Object -Property "Management Agent", "Total Objects", "Connectors", "Last Run" | Format-Table -AutoSize | Out-String -Width $Width
    New-PoshBotCardResponse -Type Normal -Text $o 
}
$CommandsToExport += 'Get-MVStats'

function Find-MVPerson {
    <#
    .SYNOPSIS
        Find User from MIM MetaVerse
    .EXAMPLE
        !MVFind
    .EXAMPLE
        !FindMV-User
    .DESCRIPTION !FindMV
        MIM MV User
    .LINK
        https://blog.darrenjrobinson.com
    #>

    [cmdletbinding()]
    [PoshBot.BotCommand(
        CommandName = 'MVFind',
        Aliases = ('FindMV-User','FindMV'),
        Permissions = 'read'
    )]
    param(
        # !MVUser <DisplayName>
        [parameter(position = 1,
            parametersetname = 'id', mandatory)]
        [string]$Identity,

        [PoshBot.FromConfig('MIMSyncCreds')]
        [parameter(mandatory)]
        [PSCredential]$MIMSyncCreds,

        [PoshBot.FromConfig('MIMSyncAddress')]
        [parameter(mandatory)]
        [string]$MIMSyncAddress
    )

    $scriptblock = {
        param($i)
        Import-Module lithnetmiisautomation 
        $query = @()
        $query += New-MVQuery -Attribute displayName -Operator Contains -Value "$($i)"
        $results = Get-MVObject -Queries $query -ObjectType person 

        if ($results.Count -gt 0) {
            # Person Output Attributes
            $MVPersonTemplate = [pscustomobject][ordered]@{ 
                samAccountName = $null 
                DisplayName    = $null 
                JobTitle       = $null
                Location       = $null              
            } 

            $resultList = @() 
            foreach ($person in $results) {
                $attributes = $person | Select-Object -Property Attributes | Select-Object -expand *
                $obj = $MVPersonTemplate.PsObject.Copy()              
                $obj.samAccountName = $attributes.samAccountName.Values.valueString
                $obj.DisplayName = $attributes.DisplayName.Values.valueString
                $obj.Location = $attributes.l.Values.valueString
                $obj.JobTitle = $attributes.JobTitle.Values.valueString
                $resultList += $obj 
            }      
            # Output
            $resultList 
        }
        else {
            Write-Output "No users containing `'$i`' found!"
        }
    }   

    # Connect to MIM Sync Server using Remote PowerShell to run LithnetMIISAutomation cmdlets  
    $result = invoke-command -ScriptBlock $scriptblock -ComputerName $MIMSyncAddress -Authentication Kerberos -Credential $MIMSyncCreds -ArgumentList $Identity
    $o = $result | Select-Object -Property samAccountName, DisplayName, Location, JobTitle | Format-Table -AutoSize | Out-String -Width $Width
    New-PoshBotCardResponse -Type Normal -Text $o 
}
$CommandsToExport += 'Find-MVPerson'

function Get-MVPerson {
    <#
    .SYNOPSIS
        Get User from MIM MetaVerse
    .EXAMPLE
        !MVUser
    .EXAMPLE
        !MVPerson
    .DESCRIPTION !MVUser
        MIM MV User
    .LINK
        https://blog.darrenjrobinson.com
    #>
    [cmdletbinding()]
    [PoshBot.BotCommand(
        CommandName = 'MVUser',
        Aliases = ('Get-MVUser', 'MVPerson', 'GetMVUser'),
        Permissions = 'read'
    )]
    param(
        # !MVUser <DisplayName>
        [parameter(position = 1,
            parametersetname = 'id', mandatory)]
        [string]$Identity,

        [PoshBot.FromConfig('MIMSyncCreds')]
        [parameter(mandatory)]
        [PSCredential]$MIMSyncCreds,

        [PoshBot.FromConfig('MIMSyncAddress')]
        [parameter(mandatory)]
        [string]$MIMSyncAddress
    )

    $scriptblock = {
        param($i)
        Import-Module lithnetmiisautomation 
        $queries = @()
        $queries += New-MVQuery -Attribute displayName -Operator Equals -Value $i
        $result = Get-MVObject -Queries $queries -ObjectType person 

        if ($result) {
            $attributes = $result | Select-Object -Property Attributes | Select-Object -expand *           

            $obj = @()
            foreach ($attr in $attributes.Keys) {
                try { 
                    # First try expanding the attribute in case it is multivalued using a comma as a separator 
                    $val = ($attributes.$attr.Values).Valuestring -join ', ' 
                } 
                catch { 
                    # Otherwise we'll just take the string value as we're outputting to strings anyway  
                    $val = $attributes.$attr.Values.Valuestring 
                } 

                # Output MV Attributes for the report 
                $mvattr = New-Object -TypeName PSObject 
                $mvattr | Add-Member -Type NoteProperty -Name Attribute -Value $attr 
                $mvattr | Add-Member -Type NoteProperty -Name Value -Value $val  
                
                $obj += $mvattr               
            }
            $obj
        }
        else {
            Write-Output "`'$i`' not found!"
        }
    }

    # Connect to MIM Sync Server using Remote PowerShell to run LithnetMIISAutomation cmdlets 
    $result = invoke-command -ScriptBlock $scriptblock -ComputerName $MIMSyncAddress -Authentication Kerberos -Credential $MIMSyncCreds -ArgumentList $Identity
    $o = $result | Select-Object -Property Attribute, Value | Sort-Object -Property Attribute  | Format-Table | Out-String -width $Width 
    New-PoshBotCardResponse -Type Normal -Text $o 
}

$CommandsToExport += 'Get-MVPerson'

function Get-MVPersonConnectors {
    <#
    .SYNOPSIS
        Get Users Connectors from MIM MetaVerse
    .EXAMPLE
        !MVUserConnectors
    .EXAMPLE
        !MVPersonConnectors
    .DESCRIPTION !MVUserConnectors
        MIM MV User Connectors
    .LINK
        https://blog.darrenjrobinson.com
    #>
    [cmdletbinding()]
    [PoshBot.BotCommand(
        CommandName = 'MVUserConnectors',
        Aliases = ('Get-MVUserConnectors', 'MVPersonConnectors', 'MVConnectors'),
        Permissions = 'read'
    )]
    param(
        # !MVUserConnectors <DisplayName>
        [parameter(position = 1,
            parametersetname = 'id', mandatory)]
        [string]$Identity,

        [PoshBot.FromConfig('MIMSyncCreds')]
        [parameter(mandatory)]
        [PSCredential]$MIMSyncCreds,

        [PoshBot.FromConfig('MIMSyncAddress')]
        [parameter(mandatory)]
        [string]$MIMSyncAddress
    )

    $scriptblock = {
        param($i)
        Import-Module lithnetmiisautomation 
        $queries = @()
        $queries += New-MVQuery -Attribute displayName -Operator Equals -Value $i
        $result = Get-MVObject -Queries $queries -ObjectType person 

        if ($result.count -eq 1) {
            $connectors = $result.CSMVLinks | Select-Object ManagementAgentName, LineageType, LineageTime 
            $connectors            
        }
        else {
            if ($result.count -eq 0){
                Write-Output "`'$i`' not found!"
            } else {
                Write-Output "`'$i`' not unique. Multiple results found. Try again to be unique."
            }
        }
    }

    # Connect to MIM Sync Server using Remote PowerShell to run LithnetMIISAutomation cmdlets 
    $result = invoke-command -ScriptBlock $scriptblock -ComputerName $MIMSyncAddress -Authentication Kerberos -Credential $MIMSyncCreds -ArgumentList $Identity
    $o = $result | Select-Object -Property ManagementAgentName, LineageType, LineageTime | Sort-Object -Property ManagementAgentName | Format-Table | Out-String -width $Width 
    New-PoshBotCardResponse -Type Normal -Text $o 
}

$CommandsToExport += 'Get-MVPersonConnectors'

Export-ModuleMember -Function $CommandsToExport