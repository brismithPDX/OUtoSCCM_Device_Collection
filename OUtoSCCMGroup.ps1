## Converts AD structure into computer collections
## Created by brian@SFTI.in with special thanks to Pasquale Lantella for providing the Convert-LHSADName function!
## 
##


## Major Global Variables and Configuration Features

##### AD OU Search Root for creating collections
#     Example OU variable fomrmat :: $OU = 'OU=Employee Computers,OU=Marylhurst Computers,DC=campus,DC=marylhurst,DC=edu'
$OU = "[SEARCH ROOT in AD FORMAT See above Example]"
#####Collection Update Frequency
$Schedule1 = New-CMSchedule -Start "01/01/2018 9:00 AM" -DayOfWeek Monday -RecurCount 1 

Function Convert-LHSADName{
    <#
    .SYNOPSIS
        Translates Active Directory names between various formats.
    
    .DESCRIPTION
        NameTranslate refers to the IADsNameTranslate interface, which can be used to convert the 
        names of Active Directory objects from one format to another.
        For Powershell we can use Microsoft´s NameTranslate COM object.
    
    .PARAMETER Identity
        Specifies an Active Directory object by providing one of the following property values. 
    
        DN                short for 'distinguished name'; e.g., 'CN=Phineas Flynn,OU=Engineers,DC=fabrikam,DC=com'
        Canonical         canonical name; e.g., 'fabrikam.com/Engineers/Phineas Flynn'
        NT4               domain\username; e.g., 'fabrikam\pflynn'
        Display           display name
        DomainSimple      simple domain name format
        EnterpriseSimple  simple enterprise name format
        GUID              GUID; e.g., '{95ee9fff-3436-11d1-b2b0-d15ae3ac8436}'
        UPN               user principal name; e.g., 'pflynn@fabrikam.com'
        CanonicalEx       extended canonical name format
        SPN               service principal name format; e.g. 'HTTP/kairomac.contoso.com'
        SID               Security Identifier; e.g., 'S-1-5-21-12986231-600641547-709122288-57999'
    
    .PARAMETER OutputType
        The output name type you want to convert to, which must be one of the following:
    
        DN                short for 'distinguished name'; e.g., 'CN=Phineas Flynn,OU=Engineers,DC=fabrikam,DC=com'
        Canonical         canonical name; e.g., 'fabrikam.com/Engineers/Phineas Flynn'
        NT4               domain\username; e.g., 'fabrikam\pflynn'
        Display           display name
        DomainSimple      simple domain name format
        EnterpriseSimple  simple enterprise name format
        GUID              GUID; e.g., '{95ee9fff-3436-11d1-b2b0-d15ae3ac8436}'
        UPN               user principal name; e.g., 'pflynn@fabrikam.com'
        CanonicalEx       extended canonical name format
        SPN               service principal name format; e.g. 'HTTP/kairomac.contoso.com'
        SID               Security Identifier; e.g., 'S-1-5-21-12986231-600641547-709122288-57999'
    
    .PARAMETER InitType
        The type of initialization to be performed, which must be one of the following:
          domain  Bind to the domain specified by the -InitName parameter
          server  Bind to the server specified by the -InitName parameter
          GC      Locate and bind to a global catalog
        
        The default value for this parameter is 'GC'. 
        When -InitType is not 'GC', you must also specify the -InitName parameter.
    
    .PARAMETER InitName
        When -InitType is 'domain' or 'server', this parameter specifies which domain or server to bind to. 
        This parameter is ignored if -InitType is 'GC'.
    
    .PARAMETER ChaseReferrals
        Switch Options of referral chasing as defined in ADS_CHASE_REFERRALS_ENUM. When referral chasing is 
        set, the name translation is performed on objects that do not belong to this directory and are 
        the referrals returned from referral chasing.
        The referral chasing options apply only when you use IADsNameTranslate::Set and IADsNameTranslate::Get.
    
    .PARAMETER Credential
        Optional, to use alternative Credentials.
    
    
    .EXAMPLE
        Get-ADComputer -Identity Server1 | Convert-LHSADName -OutputType Canonical
    
        contoso.com/Servers/Server1
    
        Get computerObject using cmdlet Get-ADComputer and pipeinput to Convert-LHSADName and convert as Type 'Canonical'
    
    .EXAMPLE
        Get-ADGroup -Identity MyGroup | Convert-LHSADName -OutputType dn
    
        CN=MyGroup,OU=Groups,DC=contoso,DC=com
    
        Convert as distinguishedName
    
    .EXAMPLE
        Convert-LHSADName -Identity 'S-1-5-21-12986231-600641999-70912299-57999' -OutputType Display
    
        John Wayne
    
        Converts a SID into a Display Name
    
    .EXAMPLE
        Convert-LHSADName -Identity TestUser@contoso.com -OutputType NT4
    
        contoso\TestUser
    
        Converts user principal name into domain\username
    
    .EXAMPLE
        $GUID = '{95ee9fff-3436-11d1-b2b0-d15ae3ac8436}'
        Convert-LHSADName -Identity $GUID -OutputType dn -InitType Domain -InitName contoso.com -Credential (Get-Credential)
    
        Converts a GUID of an ADObject of domain contoso.com into a distinguishedName using alternative Credentials. 
    
    .EXAMPLE
        Convert-LHSADName -Identity 'HTTP/kairomac.contoso.com' -OutputType dn
        CN=KAIROMAC,OU=Workstations,OU=FI,DC=contoso,DC=com
        
        Converts a Service Principal Name into a distinguishedName     
    
    .INPUTS
        System.String, you can pipe ADObjects to this Function.
    
    .OUTPUTS
        System.String
    
    .NOTES
    
        AUTHOR  : Pasquale Lantella 
        LASTEDIT: 02.08.2016
        KEYWORDS: Translate, convert, ADObjects
        Version : 1.0
        History :
    
    .LINK
        IADsNameTranslate interface
        https://msdn.microsoft.com/en-us/library/Aa706046.aspx
    
    .LINK
        ADS_NAME_TYPE_ENUM enumeration
        https://msdn.microsoft.com/en-us/library/aa772267%28v=vs.85%29.aspx
        
    .LINK
        ADS_NAME_INITTYPE_ENUM enumeration
        https://msdn.microsoft.com/en-us/library/aa772266%28v=vs.85%29.aspx
    
    .LINK
        ChaseReferrals
        https://msdn.microsoft.com/en-us/library/aa706052.aspx
    
    #Requires -Version 2.0
    #>
       
    [cmdletbinding()]  
    
    [OutputType('System.String')] 
    
    Param(
    
        [Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True,
            HelpMessage='Active Directory Object Name you want to translate.')]
        [string]$Identity,
    
        [Parameter(Position=1,Mandatory=$True,
            HelpMessage='Specify the format used for representing distinguished names.')]
        [ValidateSet("DN","Canonical","NT4","Display","DomainSimple","EnterpriseSimple","GUID","UPN","CanonicalEx","SPN","SID")]  
        [string]$OutputType,
    
        [Parameter(Position=2,HelpMessage='Type of binding to perform on a Name Translate.')] 
        [ValidateSet("GC","Domain","Server")]   
        [String]$InitType="GC",
    
        [Parameter(Position=3,HelpMessage='Translation is performed on objects that do not belong to this directory and are the referrals returned from referral chasing')]
        [Switch]$ChaseReferrals,
    
        [Parameter(Position=4)]
        [Alias('RunAs')]
        [System.Management.Automation.Credential()]$Credential = [System.Management.Automation.PSCredential]::Empty
    
       )
    
    DynamicParam {
        if ( ($PSBoundParameters['InitType']) -and ($InitType -ne "GC") )
        {
            <#
            dynamically add a new parameter called -InitName when -InitType is not 'GC' (Global Catalog Server) 
            #>
            
            #create a new ParameterAttribute Object
            $Attribute = New-Object -TypeName System.Management.Automation.ParameterAttribute
            $Attribute.Position = 3
            $Attribute.Mandatory = $true
            $Attribute.HelpMessage = "Supply the domain name within a directory forest or a machine name of a directory server."
            
            #create an attributecollection object for the attribute we just created.
            $attributeCollection = New-Object -TypeName System.Collections.ObjectModel.Collection[System.Attribute]
            
            #add our custom attribute
            $attributeCollection.Add($Attribute)
            
            #add our paramater specifying the attribute collection
            $ParameterName = 'InitName'
            $InitName_Param = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameter($ParameterName, [string], $attributeCollection)
            
            #expose the name of our parameter
            $paramDictionary = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add($ParameterName, $InitName_Param)
            return $paramDictionary
        }
    }
    
    
    BEGIN 
    {
        Set-StrictMode -Version 2.0
        ${CmdletName} = $Pscmdlet.MyInvocation.MyCommand.Name
    
    
        If (-not($PSBoundParameters['InitName']))
        {
            $InitName = $null
        }
        else
        {
            $InitName = $PSBoundParameters.InitName
        }
        Write-Debug ("Parameters:`n{0}" -f ($PSBoundParameters | Out-String))
    
    
        #region hash tables
    
        # https://msdn.microsoft.com/en-us/library/aa772267%28v=vs.85%29.aspx
        $ADS_NAME_TYPE_ENUM = @{ 
            ADS_NAME_TYPE_1779                    = 1;  # Name format as specified in RFC 1779. For example, "CN=Jeff Smith,CN=users,DC=Fabrikam,DC=com".
            ADS_NAME_TYPE_CANONICAL               = 2;  # Canonical name format. For example, "Fabrikam.com/Users/Jeff Smith".
            ADS_NAME_TYPE_NT4                     = 3;  # Account name format used in Windows. For example, "Fabrikam\JeffSmith".
            ADS_NAME_TYPE_DISPLAY                 = 4;  # Display name format. For example, "Jeff Smith".
            ADS_NAME_TYPE_DOMAIN_SIMPLE           = 5;  # Simple domain name format. For example, "JeffSmith@Fabrikam.com".
            ADS_NAME_TYPE_ENTERPRISE_SIMPLE       = 6;  # Simple enterprise name format. For example, "JeffSmith@Fabrikam.com".
            ADS_NAME_TYPE_GUID                    = 7;  # Global Unique Identifier format. For example, "{95ee9fff-3436-11d1-b2b0-d15ae3ac8436}".
            ADS_NAME_TYPE_UNKNOWN                 = 8;  <# Unknown name type. The system will estimate the format. This element is a meaningful option 
                                                           only with the IADsNameTranslate.Set or the IADsNameTranslate.SetEx method, 
                                                           but not with the IADsNameTranslate.Get or IADsNameTranslate.GetEx method.#>
            ADS_NAME_TYPE_USER_PRINCIPAL_NAME     = 9;  # User principal name format. For example, "JeffSmith@Fabrikam.com".
            ADS_NAME_TYPE_CANONICAL_EX            = 10; # Extended canonical name format. For example, "Fabrikam.com/Users Jeff Smith".
            ADS_NAME_TYPE_SERVICE_PRINCIPAL_NAME  = 11; # Service principal name format. For example, "www/www.fabrikam.com@fabrikam.com".
            ADS_NAME_TYPE_SID_OR_SID_HISTORY_NAME = 12; <# A SID string, as defined in the Security Descriptor Definition Language (SDDL), 
                                                           for either the SID of the current object or one from the object SID history. 
                                                           For example, "O:AOG:DAD:(A;;RPWPCCDCLCSWRCWDWOGA;;;S-1-0-0)" For more information, 
                                                           see Security Descriptor String Format.#>
        }
    
        # https://msdn.microsoft.com/en-us/library/aa772266%28v=vs.85%29.aspx
        $ADS_NAME_INITTYPE_ENUM = @{ 
            ADS_NAME_INITTYPE_DOMAIN = 1; # Initializes a NameTranslate object by setting the domain that the object binds to.
            ADS_NAME_INITTYPE_SERVER = 2; # Initializes a NameTranslate object by setting the server that the object binds to.
            ADS_NAME_INITTYPE_GC     = 3; # Initializes a NameTranslate object by locating the global catalog that the object binds to.
        } 
        
        # https://msdn.microsoft.com/en-us/library/aa772250.aspx
        $ADS_CHASE_REFERRALS_ENUM = @{ 
            ADS_CHASE_REFERRALS_NEVER       = (0x00); #The client should never chase the referred-to server. Setting this option prevents a client from contacting other servers in a referral process.
            ADS_CHASE_REFERRALS_SUBORDINATE = (0x20); #The client chases only subordinate referrals which are a subordinate naming context in a directory tree. For example, if the base search is requested for "DC=Fabrikam,DC=Com", and the server returns a result set and a referral of "DC=Sales,DC=Fabrikam,DC=Com" on the AdbSales server, the client can contact the AdbSales server to continue the search. The ADSI LDAP provider always turns off this flag for paged searches.
            ADS_CHASE_REFERRALS_EXTERNAL    = (0x40); #The client chases external referrals. For example, a client requests server A to perform a search for "DC=Fabrikam,DC=Com". However, server A does not contain the object, but knows that an independent server, B, owns it. It then refers the client to server B.
            ADS_CHASE_REFERRALS_ALWAYS      = (0x60); #Referrals are chased for either the subordinate or external type.
        }   
        
        #endregion hash tables
    
    
        $ADS_InitType = switch($InitType)
            {
                'Domain' {$ADS_NAME_INITTYPE_ENUM.ADS_NAME_INITTYPE_DOMAIN}
                'Server' {$ADS_NAME_INITTYPE_ENUM.ADS_NAME_INITTYPE_SERVER}
                'GC'     {$ADS_NAME_INITTYPE_ENUM.ADS_NAME_INITTYPE_GC}
                default  {$ADS_NAME_INITTYPE_ENUM.ADS_NAME_INITTYPE_GC}
            }    
    
        $ADS_OutputType = switch($OutputType)
            {
                "DN"               {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_1779}
                "Canonical"        {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_CANONICAL}
                "NT4"              {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_NT4}
                "Display"          {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_DISPLAY}
                "DomainSimple"     {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_DOMAIN_SIMPLE}
                "EnterpriseSimple" {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_ENTERPRISE_SIMPLE}
                "GUID"             {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_GUID}
                "UPN"              {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_USER_PRINCIPAL_NAME}
                "CanonicalEx"      {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_CANONICAL_EX}
                "SPN"              {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_SERVICE_PRINCIPAL_NAME}
                "SID"              {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_SID_OR_SID_HISTORY_NAME}
                "Unkonwn"          {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_UNKNOWN}
                default            {$ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_UNKNOWN}
            }
    
    
        #region Functions
    
        # Accessor functions from Bill Stewart to simplify calls to NameTranslate
        function Invoke-Method([__ComObject] $object, [String] $method, $parameters) 
        {
            $output = $Null
            $output = $object.GetType().InvokeMember($method, "InvokeMethod", $NULL, $object, $parameters)
            Write-Output $output
        }
    
        function Get-Property([__ComObject] $object, [String] $property) 
        {
            $object.GetType().InvokeMember($property, "GetProperty", $NULL, $object, $NULL)
        }
    
        function Set-Property([__ComObject] $object, [String] $property, $parameters) 
        {
            [Void] $object.GetType().InvokeMember($property, "SetProperty", $NULL, $object, $parameters)
        }
    
        #endregion Functions
    
    
    } # end BEGIN
    
    PROCESS 
    {
        #region Initialize IADsNameTranslate
        $NameTranslate = New-Object -ComObject NameTranslate
    
        If ($PSBoundParameters['Credential'])
        {
            Try
            {
                $Cred = $Credential.GetNetworkCredential()
     
                Invoke-Method $NameTranslate "InitEx" (
                    $ADS_InitType,
                    $InitName,
                    $Cred.UserName,
                    $Cred.Domain,
                    $Cred.Password
                )
            }
            Catch [System.Management.Automation.MethodInvocationException] 
            {
                Write-Error $_
                break
            }
            Finally 
            {
                Remove-Variable Cred
            }
        }
        Else
        {
            Try 
            {
                Invoke-Method $NameTranslate "Init" (
                    $ADS_InitType,
                    $InitName
                )
            }
            Catch [System.Management.Automation.MethodInvocationException] 
            {
                Write-Error $_
                break
            }
        }
        #endregion Initialize IADsNameTranslate
    
    
        If ($PSBoundParameters['ChaseReferrals']) 
        {
            Set-Property $NameTranslate "ChaseReferral" ($ADS_CHASE_REFERRALS_ENUM.ADS_CHASE_REFERRALS_ALWAYS)
        }
    
    
        Try
        {
            Invoke-Method $NameTranslate "Set" ($ADS_NAME_TYPE_ENUM.ADS_NAME_TYPE_UNKNOWN,$Identity)
            Invoke-Method $NameTranslate "Get" ($ADS_OutputType)
        }
        Catch [System.Management.Automation.MethodInvocationException] 
        {
            Write-Error "'$Identity' - $($_.Exception.InnerException.Message)"
        }
    
    } # end PROCESS
    
    END { Write-Verbose "Function ${CmdletName} finished." }
    
} ## Barrowed Function from microsoft to do Name Conversion




$SCCM_COLLECTIONS = @()
$SCCM_GROUP = Get-ADOrganizationalUnit -SearchBase $OU -SearchScope Subtree -Filter *

foreach($object in $SCCM_GROUP){
    $objectpath = $object | Convert-LHSADName -OutputType Canonical

    $SCCM_COLLECTIONS += [PSCustomObject]@{"Name" = $object.name; "Path" = $objectpath.GetValue(2)}
}

#Create SystemCenter Device Groups for each


$Query = "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System where SMS_R_System.SystemOUName ="

ForEach($Department in $SCCM_COLLECTIONS){

    $DepartmentQuery = $Query + "`"" + $Department.Path.ToUpper() + "`""

    try{
        New-CMDeviceCollection -Name $Department.Name -LimitingCollectionName "All Systems" -RefreshSchedule $Schedule1 -RefreshType Periodic
    } 
    catch [System.ArgumentException] {
        Write-Host "INFO: " + $Department.Name + " Collection already exists"
    }

    try{
        $QueryTitle = $Department.Name + " Automaticaly Generated Query"
        Add-CMDeviceCollectionQueryMembershipRule -CollectionName $Department.Name -QueryExpression $DepartmentQuery -RuleName $QueryTitle
    } 
    catch [System.ArgumentException] {
        Write-Host "INFO: " + $Department.Name + " Rule already exists"
    }
}

Write-Host "========================================="
Write-Host "========================================="
Write-Host "==  SCCM Device Collections Created    =="]
Write-Host "== Special Thanks to Pasquale Lantella =="
Write-Host "==      and Marylhurst University      =="
Write-Host "========================================="
Write-Host "========================================="