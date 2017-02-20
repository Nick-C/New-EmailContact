Function New-EmailContact {

    <#

    .SYNOPSIS

    Takes either a CSV file or manually specified parameters and creates an Exchange Mail Contact and sets Send As permissions for the specified user.

    .DESCRIPTION

    The New-EmailMask function takes either a path to a csv file containing one or more masks to setup or commandline parameters to create a single mask. The function connects to
    an Exchange 2013 server via remote powershell and creates the Mail Contact(s) and then grants the specified user Send As permissions on those contacts.

    .PARAMETER filePath

    A path to a csv file that contains fields labelled externalName, externalEmail, userEmail.

    .PARAMETER externalName

    The External Name to be used for the Email mask.

    .PARAMETER externalEmail

    The email address to be used for the email mask (this is the email address provided by the client).

    .PARAMETER userEmail

    The email address of the internal user who will be using the mask.

    .PARAMETER domainController

    IP or FQDN of the domain controller to bind to (optional).

    .PARAMETER exchangeServer

    URL to the Powershell folder of an Exchange 2013 server in the format http://server.domain.com/PowerShell (optional).

    .PARAMETER searchBase

    LDAP query to specify where to search for AD user that matches the userEmail field (optional).

    .PARAMETER Credentials

    Set this to be prompted to provide specific credentials to connect to Exchange otherwise the current logged in user will be used - user the domain\username format. (optional).

    .EXAMPLE

    Create multiple email masks from a CSV file.

    New-EmailMask -filePath c:\temp\emailmasks.csv

    .EXAMPLE

    Create a single email mask by passing command line parameters

    New-EmailMask -externalName Test -externalEmail Test@test.com -userEmail joebloggs@domain.com

    .EXAMPLE

    Create a single email mask by passing command line parameters and specify different credentials to use when connecting to Exchange.

    New-EmailMask -externalName Test -externalEmail Test@test.com -userEmail joebloggs@domain.com -Credentials

    .NOTES

    To run this cmdlet you need access to roles on Exchange to both create new Mail Contact objects and set AD Permissions, if your current user doesn't have this access specify the -Credentials 
    parameter to be prompted for additional login details.

    #>

    [CmdletBinding(DefaultParameterSetName = 'manual')]

    Param(
    [Parameter(Mandatory = $true, ParameterSetName = 'filePath')][string]$filePath,
    [Parameter(Mandatory = $true, ParameterSetName = 'manual')][string]$externalName,
    [Parameter(Mandatory = $true, ParameterSetName = 'manual')][string]$externalEmail,
    [Parameter(Mandatory = $true, ParameterSetName = 'manual')][string]$userEmail,
    [string]$domainController,
    [string]$exchangeServer,
    [string]$searchBase,
    [switch]$Credentials
    )

    $status = 0

    # Disable Debugging by default
    #$null = Set-PSDebug -Off

    # Suppress Debug prompts
    #$DebugPreference = "Continue"

    # Make all errors hard stopping for catching Exchange commandlet errors.
    $global:ErrorActionPreference = "Stop"; 

    # Error handling
    trap {
    	write-host "SCRIPT EXCEPTION: $($_.Exception.Message)";
        if($Session){
           Remove-PSSession $Session
           }
    	#exit 2;
        }

    # If no DC specified lets set one now so all commands go through same place
    if(!$domainController){
        $domainController = "your domain controller here"
        }

    # Lets also set a searchbase ready for AD queries
    if (!$searchBase){
       $searchBase = "dc=yourdomain,dc=co,dc=uk" 
       }

    # And lets set the Exchange server if we have had one specified
    if (!$exchangeServer){
       $exchangeServer = "http://exchange2013.yourdomain.com/PowerShell"
       }

    # For debugging, lets just confirm what all variables are set to
    Write-Verbose "filePath:  $($filePath)"
    Write-Verbose "externalName: $($externalName)"
    Write-Verbose "externalEmail: $($externalEmail)"
    Write-Verbose "userEmail: $($userEmail)"
    Write-Verbose "domainController: $($domainController)"
    Write-Verbose "exchangeServer: $($exchangeServer)"
    Write-Verbose "searchBase: $($searchBase)"


    #Setup and Connect to remote Exchange PS Session

    Write-Verbose "Connecting to Remote Powershell on $($exchangeServer)"

    # If -Credentials flag set on cmd then lets prompt to get them before connecting to exchange, if not then carry on connecting with current creds
    if($Credentials){
        $Session = New-PSSession –ConfigurationName Microsoft.Exchange –ConnectionUri $exchangeServer -Authentication Kerberos -Credential (Get-Credential)
        } else {
        $Session = New-PSSession –ConfigurationName Microsoft.Exchange –ConnectionUri $exchangeServer -Authentication Kerberos
        }
    $null = Import-PSSession -Session $Session -DisableNameChecking

    # Add AD module

    $null = Import-Module ActiveDirectory

    # If no csv file provided, pull details from other params
    if (!$filePath){
        Write-Verbose "No CSV file specified via -filepath so creating single contact"
               
        # Search AD for user based on $userEmail
        Write-Verbose "Searching AD for a user with email address $($userEmail)"
        $user = Get-ADUser -Filter {EmailAddress -eq $userEmail} -SearchBase $searchBase -Server $domainController -ErrorAction Stop
                
        Write-Verbose "User selected: $($user.Name)"
    
        # Now we have the users details lets start by creating the Exchange contact
        $contactName = $externalName + ' - ' + $user.Name
        Write-Verbose "Contact Name will be $($contactName)"

        # Need to trim any spaces in externalName before using it for the alias or else Exchange will error
        $trimmedexternalName = $externalName -replace(" ","")

        # Exchange Alias can't be longer than 64 chars so lets check length and trim if needed
        if ($trimmedexternalName.Length -gt 64) { $trimmedexternalName.Substring(0, $length) }
        Write-Verbose "Trimmed externalName is: $($trimmedexternalName)"

        # Now we have the full name for the new contact lets get Exchange involved and create it
        $null = New-MailContact -Name $contactName -FirstName $user.GivenName -LastName $user.Surname -Alias "$($user.SamAccountName)$($trimmedexternalName)"-ExternalEmailAddress $externalEmail -OrganizationalUnit "Alternative Emails" -DomainController $domainController

        # And now lets get the new contact so we can set permissions on it
        Write-Verbose "Setting permissions on contact"
        $mailContact = Get-MailContact -Identity $contactName -DomainController $domainController
        $mailContact | Write-Verbose 

        # And lets grant send as permissions on the contact to the user
        $null = Add-ADPermission -Identity $contactName -User $user.SamAccountName -AccessRights ExtendedRight -ExtendedRights "Send As" -DomainController $domainController

        Write-Host "Created new contact named $($contactName) and granted Send As rights on it to the user $($user.Name)"
                
        }
        else {
        # If $filepath exists lets check if it is a valid file path
        $ValidPath = Test-Path $filePath -IsValid
        if($ValidPath -eq $false){
            throw "Path to CSV File is invalid or does not exist"
        }

        # If we made it this far then assume $filepath is ok and try to import it
        $masks = Import-Csv $filePath

        ForEach($emailmask in $masks){

            $externalName = "$($emailmask.externalName)"
            $externalEmail = "$($emailmask.externalEmail)"
            $userEmail = "$($emailmask.userEmail)"
            Write-Verbose "externalName: $($externalName)"
            Write-Verbose "externalEmail: $($externalEmail)"
            Write-Verbose "userEmail: $($userEmail)"

            # Search AD for user based on $userEmail
            Write-Verbose "Searching AD for a user with email address $($userEmail)"
            $user = Get-ADUser -Filter {EmailAddress -eq $userEmail} -SearchBase $searchBase -Server $domainController -ErrorAction Stop
                
            Write-Verbose "User selected: $($user.Name)"
    
            # Now we have the users details lets start by creating the Exchange contact
            $contactName = $externalName + " - " + $user.Name
            Write-Verbose "Contact Name will be $($contactName)"

            # Need to trim any spaces in externalName before using it for the alias or else Exchange will error
            $trimmedexternalName = $externalName -replace(" ","")

            # Exchange Alias can't be longer than 64 chars so lets check length and trim if needed
            if ($trimmedexternalName.Length -gt 64) { $trimmedexternalName.Substring(0, $length) }

            Write-Verbose "Trimmed externalName is: $($trimmedexternalName)"

            # Now we have the full name for the new contact lets get Exchange involved and create it
            $null = New-MailContact -Name $contactName -FirstName $user.GivenName -LastName $user.Surname -Alias "$($user.SamAccountName)$($trimmedexternalName)"-ExternalEmailAddress "$($externalEmail)" -OrganizationalUnit "Alternative Emails" -DomainController $domainController

            # And now lets get the new contact so we can set permissions on it
            Write-Verbose "Setting permissions on contact"
            $mailContact = Get-MailContact -Identity $contactName -DomainController $domainController
            $mailContact | Write-Verbose 

            # And lets grant send as permissions on the contact to the user
            $null = Add-ADPermission -Identity $contactName -User $user.SamAccountName -AccessRights ExtendedRight -ExtendedRights "Send As" -DomainController $domainController

            Write-Host "Created new contact named $($contactName) and granted Send As rights on it to the user $($user.Name)"
            }

        }

        # Clean up and close down the Remote Exchange PS Session
        Remove-PSSession $Session
}