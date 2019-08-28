@{
    RootModule = 'PoshBot.LithnetRMA.psm1'
    ModuleVersion = '1.0.0'
    
    Description = 'PoshBot module for the Lithnet RMA PowerShell Module'
    Author = 'Darren J Robinson'
    CompanyName = 'Community'
    Copyright = '(c) 2019 Darren J Robinson. All rights reserved.'
    PowerShellVersion = '5.0.0'
    
    GUID = '587e7810-f876-40b1-b33b-c58d20e143b9'
    
    RequiredModules = @('PoshBot')
    FunctionsToExport = '*'
    
    PrivateData = @{
        # These are permissions we'll expose in our poshbot module even though version 1.0.0 only provides Read Functions
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
