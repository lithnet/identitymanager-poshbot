Import-Module LithnetRMA

# Slack text width with the formatting we use maxes out ~80 characters...
$Width = 80
$CommandsToExport = @()

function Find-Person {
    <#
    .SYNOPSIS
        Find MIM Resource (Person)
    .PARAMETER Identity <AccountName>|<DisplayName>
        MIM Service AccountName or DisplayName
    .USAGE
        !Find <AccountName>|<DisplayName>
    .EXAMPLE
        !Find 'Darren Robinson'        
    .EXAMPLE
        !Search 'Darren Robinson'
    .Example
        !Find drobinson
    .Example
        !Search drobinson
    .DESCRIPTION !Find
        Search the MIM Service for Person objects using AccountName or DisplayName
    .LINK
        https://blog.darrenjrobinson.com
    #>
    [cmdletbinding()]
    [PoshBot.BotCommand(
        CommandName = 'Find',
        Aliases = ('Search', 'Search-Resources'),
        Permissions = 'read'
    )]
    param(
        # !Find <AccountName>|<DisplayName>        
        [parameter(position = 1,
            parametersetname = 'id', mandatory)]
        [string]$Identity,

        [PoshBot.FromConfig('MIMServiceCreds')]
        [parameter(mandatory)]
        [PSCredential]$MIMServiceCreds,

        [PoshBot.FromConfig('MIMServiceAddress')]
        [parameter(mandatory)]
        [string]$MIMServiceAddress
    )

    # Person Output Attributes
    $MIMPersonTemplate = [pscustomobject][ordered]@{ 
        AccountName = $null 
        DisplayName = $null 
        JobTitle    = $null
        City        = $null              
    } 

    # Connect to MIM Service using LithnetRMA PSModule using configuration info from the PoshBot Config   
    Set-ResourceManagementClient -BaseAddress $MIMServiceAddress -Credentials $MIMServiceCreds
    # Try to find the user using DisplayName
    $result = Search-Resources -Xpath "/Person[starts-with(DisplayName, `'$Identity`')]" -AttributesToGet @("AccountName", "DisplayName", "JobTitle", "City", "EmployeeType", "Email")

    if (!$result.AccountName) {        
        # Failed finding user by DisplayName, try to find the user using AccountName 
        $result = Search-Resources -Xpath "/Person[starts-with(AccountName, `'$Identity`')]" -AttributesToGet @("AccountName", "DisplayName", "JobTitle", "City", "EmployeeType", "Email")
    }

    if ($result.AccountName) {
        $resultList = @()
        foreach ($person in $result) {
            $o = $MIMPersonTemplate.PsObject.Copy()
            $o.AccountName = $person.AccountName
            $o.DisplayName = $person.DisplayName
            $o.City = $person.City
            $o.JobTitle = $person.JobTitle 
            
            $resultList += $o 
        }      
        $o = $resultList | Format-Table -AutoSize | Out-String -Width $Width
    }
    else {
        $o = "User `'$($Identity)`' not found!"
    }

    New-PoshBotCardResponse -Type Normal -Text $o 
}
$CommandsToExport += 'Find-Person'

function Get-Person {
    <#
    .SYNOPSIS
        Get MIM Resource
    .PARAMETER Identity <AccountName>|<DisplayName>
        MIM Service AccountName or DisplayName
    .USAGE
        !Person <AccountName>|<DisplayName>
    .EXAMPLE
        !Person 'Darren Robinson'        
    .EXAMPLE
        !User 'Darren Robinson'
    .Example
        !Person drobinson
    .DESCRIPTION !Person
        Get a Person Object from the MIM Service using AccountName or DisplayName
    .LINK
        https://blog.darrenjrobinson.com
    #>
    [cmdletbinding()]
    [PoshBot.BotCommand(
        CommandName = 'Person',
        Aliases = ('User', 'Get-Resource','GetUser', 'GetPerson'),
        Permissions = 'read'
    )]
    param(
        # !Person <AccountName>|<DisplayName>
        [parameter(position = 1,
            parametersetname = 'id')]
        [string]$Identity,

        [PoshBot.FromConfig('MIMServiceCreds')]
        [parameter(mandatory)]
        [PSCredential]$MIMServiceCreds,

        [PoshBot.FromConfig('MIMServiceAddress')]
        [parameter(mandatory)]
        [string]$MIMServiceAddress
    )

    # Person Output Attributes
    $MIMPersonTemplate = [pscustomobject][ordered]@{ 
        AccountName  = $null 
        DisplayName  = $null 
        EmployeeType = $null 
        JobTitle     = $null
        City         = $null   
        Email        = $null                 
    } 

    # Connect to MIM Service using LithnetRMA PSModule using configuration info form the PoshBot Config   
    Set-ResourceManagementClient -BaseAddress $MIMServiceAddress -Credentials $MIMServiceCreds

    try {
        # Try to find the user using DisplayName
        $result = Get-Resource -AttributeName DisplayName -AttributeValue "$($Identity)" -ObjectType Person 
    }
    catch {
        $result = Get-Resource -AttributeName AccountName -AttributeValue "$($Identity)" -ObjectType Person 
    }
    
    if ($result.AccountName) {
        $o = $MIMPersonTemplate.PsObject.Copy()
        $o.AccountName = $result.AccountName
        $o.DisplayName = $result.DisplayName
        $o.City = $result.City
        $o.JobTitle = $result.JobTitle    
        $o.EmployeeType = $result.EmployeeType
        $o.Email = $result.Email                    
        $o = $o | Format-List | Out-String -width $Width       
    }
    else {
        $o = "`'$($Identity)`' not found. Try Searching for the user using !Find <accountname>|<fullname>"
    }
    
    New-PoshBotCardResponse -Type Normal -Text $o 
}
$CommandsToExport += 'Get-Person'

Export-ModuleMember -Function $CommandsToExport