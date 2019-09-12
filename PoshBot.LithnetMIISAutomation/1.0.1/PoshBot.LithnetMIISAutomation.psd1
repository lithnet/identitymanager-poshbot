@{
    RootModule = 'PoshBot.LithnetMIISAutomation'   
    ModuleVersion = '1.0.1'
    
    Description = 'PoshBot module for the Lithnet MIISAutomation PowerShell Module'
    Author = 'Darren J Robinson'
    CompanyName = 'Community'
    Copyright = '(c) 2019 Darren J Robinson. All rights reserved.'
    PowerShellVersion = '5.0.0'
    
    GUID = '587e7810-f876-40b1-b33b-c58d20e143a0'
    RequiredModules = @('PoshBot')
    FunctionsToExport = '*'
    
    PrivateData = @{
        # Both Read and Write specified, but all current functions are Read only.
        Permissions = @(
            @{
                Name = 'read'
                Description = 'Run commands that have Read Only access to the MIM Service'
            }
            @{
                Name = 'write'
                Description = 'Run commands that have Write access to the MIM Service'
            }
        )
    } # End of PrivateData hashtable
}
