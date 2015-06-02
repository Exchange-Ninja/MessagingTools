Function Get-MessageCount {
<#
.Synopsis
    Get-MessageCount cmdlet is used to provide either an overview or detailed description of message sent or received for a server or servers.

.DESCRIPTION
    Get-MessageCount cmdlet is used to provide either an overview or detailed description of message sent or received for a server or servers. Get-MessageCount is designed to display received messages by default. By specifying the Send parameter it will also display an overview for the sent email. Get-MessageCount is also designed to pull data 24 hours from 20:00:00 this can be changed by opting to use the Day and 
    time parameter. 
.EXAMPLE 
    Get-MessageCount -Server EX01


    Type                      : Overview - Receive
    Date Range                : 05/27/2015 20:00:00 to 05/28/2015 16:17:07
    Server                    : EX01
    Message Count             : 1203
    Total Message Size (MB)   : 324.8
    Average Message Size (KB) : 276.47

    This example shows just specifying a server and being provided an overview count for the server.
.EXAMPLE
    Get-MessageCount -Server EX01, EX02


    Type                      : Overview - Receive
    Date Range                : 05/27/2015 20:00:00 to 05/28/2015 16:17:42
    Server                    : EX01
    Message Count             : 1203
    Total Message Size (MB)   : 324.8
    Average Message Size (KB) : 276.47

    Type                      : Overview - Receive
    Date Range                : 05/27/2015 20:00:00 to 05/28/2015 16:17:42
    Server                    : EX02
    Message Count             : 1213
    Total Message Size (MB)   : 329.45
    Average Message Size (KB) : 278.11

    This example shows how to specify multiple servers.
.EXAMPLE
    Get-MessageCount -Server EX01, EX02 -Day 5 -Time 18:00:00 -Send


    Type                      : Overview - Receive
    Date Range                : 05/23/2015 18:00:00 to 05/28/2015 16:19:34
    Server                    : EX01
    Message Count             : 5815
    Total Message Size (MB)   : 1253.09
    Average Message Size (KB) : 220.66

    Type                      : Overview - Send
    Date Range                : 05/23/2015 18:00:00 to 05/28/2015 16:19:34
    Server                    : EX01
    Message Count             : 1472
    Total Message Size (MB)   : 83.66
    Average Message Size (KB) : 58.2

    Type                      : Overview - Receive
    Date Range                : 05/23/2015 18:00:00 to 05/28/2015 16:19:34
    Server                    : EX02
    Message Count             : 5624
    Total Message Size (MB)   : 1184.36
    Average Message Size (KB) : 215.65

    Type                      : Overview - Send
    Date Range                : 05/23/2015 18:00:00 to 05/28/2015 16:19:34
    Server                    : EX02
    Message Count             : 1472
    Total Message Size (MB)   : 83.81
    Average Message Size (KB) : 58.3

    This example shows requesting multiple servers and specifying 2 days, starting at 18:00:00 and requesting Send overview.
.EXAMPLE       
    Get-MessageCount -Server EX01 -Send -Detail | Where-Object {$_.'Message Count' -ge 1}
    

    Type                      : Detail - Receive
    Date Range                : 05/27/2015 20:00:00 to 05/28/2015 16:18:06
    Server                    : EX01
    Receive Connector ID      : Default EX01
    Message Count             : 1040
    Total Message Size (MB)   : 312.39
    Average Message Size (KB) : 307.59

    Type                          : Detail - Send
    Date Range                    : 05/27/2015 20:00:00 to 05/28/2015 16:18:06
    Server                        : EX01
    Send Connector ID             : Intra-Organization SMTP Send Connector
    Message Count                 : 268
    Average Message Latency (Sec) : 2.68
    Average Message Size (KB)     : 75.81
    Total Message Size (MB)       : 19.84

    This example is using the Where-Object cmdlet to only provide the Receive and Send connectors that have a message count greater than or eqal to 1 when specifying the Detail parameter.
#>
    [CmdletBinding()]
    Param
    (
        # The Server parameter specifies the Exchange 2010 server that contains the message tracking logs to be searched. 
        # The Server parameter can take any of the following values for the target server:
        # * Name
        # * FQDN
        # * Distinguished name (DN)
        # * Legacy Exchange DN
        # * GUID
        [Parameter( Mandatory=$True,
                    ValueFromPipelineByPropertyName=$True)]
        [Alias( 'HostName','ComputerName','ServerName','Name' )]
        [String[]]
        $Server,

        # The Day parameter is used to return message tracking log entries the number of days specified back from when the cmdlet is ran.
        # By default the day parameter is set to 1 day. 
        [Parameter(Mandatory=$False)]
        [Int32]
        $Day = '1',

        # The Time parameter is used to return message tracking log entries between the specified time. Entries are returned up 
        # to the specified time. The time must be specified in the format hh:mm:ss. 
        # 12:00:00 returns results from the previous day at 12:00:00.
        [Parameter(Mandatory=$false)]
        [String]
        $Time,

        # The Detailed parameter is used to provide more detailed information in the form of providing each send and receive connectors for the specified server.
        [switch]
        $Detail,

        # The Send paramter is used to to select an overview of sent mail. If used in conjunction with the Detail parameter
        # it will provide every send connector associated with the specified server.
        [Switch]
        $Send
    )

    Begin {
        Write-Verbose "Start Begin Block"        
        Write-Verbose "Verifying if Microsoft.Exchange.Management.PowerShell.E2010 snapin has been loaded"          
            If ( ! ( Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"} ) ) {
	            try{
                    Write-Verbose "Attempting to load Microsoft.Exchange.Management.PowerShell.E2010"
		                Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
                    Write-Verbose "Microsoft.Exchange.Managegment.Powershell.E2010 snapin has been loaded"
	            }
	            catch{
		            Write-Warning $_.Exception.Message
		            EXIT
	            }
            }                    
        Write-verbose "Creating required Variables for Get-MessageCountReport cmdlet"        
            If ($Time) {
                Write-verbose "Time parameter specified for $Time applying to required variables"
                    [DateTime]$Start = (Get-Date).AddDays(-$Day).ToShortDateString() + " $Time"
                    [DateTime]$End   = (Get-Date).ToShortDateString() + " $Time"
                Write-verbose "Time parameter applied to required variables"
            }
            Else{
                [DateTime]$Start = (Get-Date).AddDays(-$Day).ToString()
                [DateTime]$End   = (Get-Date).ToString()
            }
        Write-Verbose "Completed Begin Block"
    }
    Process {
        Write-Verbose "Start Process Block"
            Write-Verbose "Start Iteration through Servers"
                ForEach ( $ServerName in $Server ) {           
                    Try { 
                        Write-Verbose "Verify if $ServerName is Valid"
                            $Valid = Get-TransportServer -Identity $ServerName -ErrorAction Stop
                        Write-Verbose "Verified $ServerName is Valid"
                    }
                    Catch {                    
                        Write-Warning $_.Exception.Message
                    }        
                    If ( $Valid ) {
                        Try {
                            Write-Verbose "Obtain Message Logs for $ServerName"
                                $MessageLogs = Get-MessageTrackingLog -Server $ServerName -Start $Start -End $End -ResultSize Unlimited -ErrorAction Stop | Select ConnectorID, EventID, TotalBytes, MessageLatency
                            Write-Verbose "Message Logs obtained for $ServerName"
                        }
                        Catch {
                            Write-Warning $_.Exception.Message                    
                        }
                        If ( $Detail ) {                            
                            Write-Verbose "Detailed Parameter selected"
                            Write-Verbose "Start Detail Process"
                                Write-Verbose "Filtering Message Logs for Receive events"
                                    $ReceiveMessageLogs = $MessageLogs | Where-Object { $_.EventID -eq "Receive" }
                                Write-Verbose "Message Logs filted for Receive events"
                                Write-Verbose "Obtain Receive Connector list for $ServerName"
                                    $ReceiveConnectorList = Get-ReceiveConnector -Server $ServerName | Select-Object -ExpandProperty Name                        
                                Write-Verbose "Receive Connector list obtained for $ServerName"
                                Write-Verbose "Iterate through Receive Connectors for $ServerName Receive Message Logs"                                                    
                                    ForEach ( $ReceiveConnector in $ReceiveConnectorList ) {
                                        Write-Verbose "Filtering Messagelogs for $ReceiveConnector"
                                            $ReceiveConnectorLogs = $ReceiveMessageLogs | Where-Object { $_.ConnectorID -eq "$ServerName\$ReceiveConnector" }                                    
                                        Write-Verbose "Object being created"                                                                              
                                            $Property = @{  'Type'                      = "Detail - Receive"
                                                            'Date Range'                = "$Start to $End"
                                                            'Server'                    = $ServerName
                                                            'Receive Connector ID'      = $ReceiveConnector
                                                            'Message Count'             = $ReceiveConnectorLogs.count
                                                            'Total Message Size (MB)'   = [System.Math]::Round( ( $ReceiveConnectorLogs | Measure-Object TotalBytes -Sum ).Sum / 1MB, 2 )
                                                            'Average Message Size (KB)' = [System.Math]::Round( ( $ReceiveConnectorLogs | Measure-Object TotalBytes -Average ).Average / 1KB, 2 )
                                                        }
                                            $Object = New-Object -TypeName PSCustomObject -Property $Property
                                        Write-Verbose "Object Created"
                                        Write-Verbose "Outputting Object Values"                                    
                                            Write-Output $Object | Select 'Type', 'Date Range', 'Server', 'Receive Connector ID', 'Message Count', 'Total Message Size (MB)', 'Average Message Size (KB)'                                        
                                        Write-Verbose "Object Values outputted"
                                        Write-Verbose "Message logs filtered for $ReceiveConnector"
                                        #Write-Verbose "Clean up variables"
                                            $ReceiveConnectorLogs = $null
                                        #Write-Verbose "Variables cleaned"
                                    }
                                Write-Verbose "Iteration through Receive Connectors complete for $ServerName"                                
                                If ( $Send ) {
                                    Write-Verbose "Send parameter selected"
                                    Write-Verbose "Filtering Message Logs for Send events"
                                        $SendMessageLogs = $MessageLogs | Where-Object { $_.EventID -eq "SEND" }
                                    Write-Verbose "Message Logs filted for Send events"
                                    Write-Verbose "Obtain Send Connector list for $ServerName"
                                        $SendConnectorList = ( Get-SendConnector | Select-Object -ExpandProperty Name ) + "Intra-Organization SMTP Send Connector"                         
                                    Write-Verbose "Send Connector list obtained for $ServerName"
                                    Write-Verbose "Iterate through Send Connectors for $ServerName Send Message Logs"                                                    
                                        ForEach ( $SendConnector in $SendConnectorList ) {                        
                                            Write-Verbose "Filtering Messagelogs for $SendConnector"
                                                $SendConnectorLogs = $SendMessageLogs | Where-Object { $_.ConnectorID -eq "$SendConnector" }                                    
                                            Write-Verbose "Object being created"                                                                              
                                                $Property = @{  'Type'                          = "Detail - Send"
                                                                'Date Range'                    = "$Start to $End"
                                                                'Server'                        = $ServerName
                                                                'Send Connector ID'             = $SendConnector
                                                                'Message Count'                 = $SendConnectorLogs.count 
                                                                'Average Message Latency (Sec)' = [System.Math]::Round( ( $SendConnectorLogs | Select-Object -ExpandProperty MessageLatency | Measure-Object -Property TotalSeconds -Average).Average / 60, 2)
                                                                'Total Message Size (MB)'       = [System.Math]::Round( ( $SendConnectorLogs | Measure-Object -Property TotalBytes -Sum ).Sum / 1MB, 2 )
                                                                'Average Message Size (KB)'     = [System.Math]::Round( ( $SendConnectorLogs | Measure-Object -Property TotalBytes -Average ).Average / 1KB, 2 )
                                                            }
                                                $Object = New-Object -TypeName PSCustomObject -Property $Property 
                                            Write-Verbose "Created Object"                                    
                                                Write-Output $Object | Select 'Type', "Date Range", 'Server', 'Send Connector ID', 'Message Count', 'Average Message Latency (Sec)', 'Average Message Size (KB)', 'Total Message Size (MB)'                                     
                                            Write-Verbose "Message logs filtered for $SendConnector"
                                            Write-Verbose "Clean up variables"
                                                $SendConnectorLogs = $null
                                            Write-Verbose "Variables cleaned"
                                        }
                                    Write-Verbose "Iteration through Send Connectors complete for $ServerName"
                                }
                            Write-Verbose "Complete Detail process"
                        }
                        Else {
                            Write-Verbose "Detailed parameter not selected"
                            Write-Verbose "Start Overview process"
                                Write-Verbose "Filtering Message Logs for Receive events"
                                    $ReceiveMessageLogs = $MessageLogs | Where-Object { $_.EventID -eq "Receive" }
                                Write-Verbose "Message Logs filted for Receive events"
                                Write-Verbose "Object being created"
                                    $Property = @{  'Type'                      = "Overview - Receive"
                                                    'Date Range'                = "$Start to $End"
                                                    'Server'                    = $ServerName
                                                    'Message Count'             = $ReceiveMessageLogs.count
                                                    'Total Message Size (MB)'   = [System.Math]::Round( ( $ReceiveMessageLogs | Measure-Object TotalBytes -Sum ).Sum / 1MB, 2 )
                                                    'Average Message Size (KB)' = [System.Math]::Round( ( $ReceiveMessageLogs | Measure-Object TotalBytes -Average ).Average / 1KB, 2 )
                                                }
                                    $Object = New-Object -TypeName PSCustomObject -Property $Property
                                Write-Verbose "Object Created"
                                Write-Verbose "Outputting Object values"                                    
                                    Write-Output $Object | Select 'Type', 'Date Range', 'Server', 'Message Count', 'Total Message Size (MB)', 'Average Message Size (KB)'                                        
                                Write-Verbose "Object values outputted"
                                Write-Verbose "Clean up variables"
                                    $ReceiveMessageLogs = $null
                                Write-Verbose "Variables cleaned"
                                If ( $Send ) {
                                    Write-Verbose "Filtering Message Logs for Send events"
                                        $SendMessageLogs = $MessageLogs | Where-Object { $_.EventID -eq "SEND" }
                                    Write-Verbose "Message Logs filted for Send events"
                                    Write-Verbose "Object being created"
                                        $Property = @{  'Type'                          = "Overview - Send"
                                                        'Date Range'                    = "$Start to $End"
                                                        'Server'                        = $ServerName
                                                        'Message Count'                 = $SendMessageLogs.count
                                                        'Average Message Latency (Sec)' = [System.Math]::Round( ( $SendMessageLogs | Select -ExpandProperty MessageLatency | Measure-Object -Property TotalSeconds -Average).Average, 2)
                                                        'Total Message Size (MB)'       = [System.Math]::Round( ( $SendMessageLogs | Measure-Object -Property TotalBytes -Sum ).Sum / 1MB, 2 )
                                                        'Average Message Size (KB)'     = [System.Math]::Round( ( $SendMessageLogs | Measure-Object -Property TotalBytes -Average ).Average / 1KB, 2 )
                                                    }
                                        $Object = New-Object -TypeName PSCustomObject -Property $Property
                                    Write-Verbose "Object Created"
                                    Write-Verbose "Outputting Object values"                                    
                                        Write-Output $Object | Select 'Type', 'Date Range', 'Server', 'Message Count', 'Average Message Latency (Sec)', 'Total Message Size (MB)', 'Average Message Size (KB)'                                        
                                    Write-Verbose "Object values outputted"
                                    Write-Verbose "Clean up variables"
                                        $SendMessageLogs = $null
                                    Write-Verbose "Variables cleaned"
                                }
                            Write-Verbose "Complete Overview process"
                        }
                    }
                    Write-Verbose "Clean up Server specific variables"
                        $MessageLogs        = $null
                        $ReceiveMessageLogs = $null
                        $SendMessageLogs    = $null
                    Write-Verbose "Server specific variables cleaned" 
                }
            Write-Verbose "Completed Iteration through Servers"
        Write-Verbose "Completed Process Block"
    }
    End {
        Write-Verbose "Start End Block"
        Write-Verbose "Completed End Block" 
    }
}
Function Get-PublicFolderReport {
<#
.Synopsis
   Get-PublicFolderReport creates a report for PublicFolder Databases. 
.DESCRIPTION
   Get-PublicFolderReport is an advanced function cmdlet that queries Public Folder Databases in Exchange 2010. If you have multiple versions of Exchange in your environment for example Exchange 2003 or Exchange 2007 you may use the -IncludePreExchange2010 switch parameter to include preExchange2010 databases.
.EXAMPLE
    Get-PublicFolderReport -Identity PUBDB1

    Server                 : EX01
    Public Folder Database : PUBDB1
    Item Count             : 665860
    Deleted Item Count     : 249
    Database Size          : 172.8 GB (185,497,354,240 bytes)
    White Space            : 43.31 MB (45,416,448 bytes)

    This examples shows how to use the cmdlet requesting a specific Public Folder Database.
.EXAMPLE
   Get-PublicFolderReport -Identity PUBDB1, PUBDB2


    Server                 : EX01
    Public Folder Database : PUBDB1
    Item Count             : 665860
    Deleted Item Count     : 249
    Database Size          : 172.8 GB (185,497,354,240 bytes)
    White Space            : 41.31 MB (43,319,296 bytes)

    Server                 : EX02
    Public Folder Database : PUBDB2
    Item Count             : 666210
    Deleted Item Count     : 361
    Database Size          : 173.1 GB (185,900,007,424 bytes)
    White Space            : 127.3 MB (133,529,600 bytes)

    This examples shows how to use the cmdlet requesting multiple Public Folder Databases.
.EXAMPE
    Get-PublicFolderReport -Server SCEA0XM21


    Server                 : SCEA0XM21
    Public Folder Database : PUBDB2
    Item Count             : 666211
    Deleted Item Count     : 361
    Database Size          : 173.1 GB (185,900,007,424 bytes)
    White Space            : 130.4 MB (136,708,096 bytes)


    This example shows how to use the cmlet requesting a all Public Folder Database from a server. 
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
    [CmdletBinding()]
    Param
    (
        # The Identity parameter specifies a public folder database. You can use the following values:
        # * GUID
        # * Distinguished name (DN)
        # * Database name
        # If you don't specify the server name, the cmdlet searches for public folder databases on the local server. If you have multiple databases with the same name, the cmdlet retrieves all
        # databases with the same name in the specified scope.
        # This parameter cannot be used with the Server parameter.
        [Parameter(Mandatory=$False,
                   ParameterSetName="Identity",
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [Alias("DatabaseName",
               "Database")] 
        [string[]]$Identity,

        # The Server parameter specifies the name of a server that contains a public folder database; by default Server paramter is LocalHost. If specified, only the public folder databases on the specified server/s ares returned.
        # This parameter cannot be used with the Identity parameter.

        [Parameter(Mandatory=$False,
                   ParameterSetName="Server",
                   ValueFromPipelineByPropertyName=$False,
                   Position=0)]
        [Alias("Computer",
               "ComputerName",
               "ServerName")]
        [String[]]$Server = 'LocalHost',

        # The IncludePreExchange2010 parameter specifies whether information about all the public folder databases in the organization is returned, including versions of Microsoft 
        # Exchange prior to Exchange Server 2010. If the IncludePreExchange2010 parameter is included in the command, information about all of the public folder databases in the 
        # organization is returned. This parameter isn't necessary when the Server or Identity parameter is included in the command.
        [Switch]$IncludePreExchange2010
    )
    Begin {
        If ( ! ( Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"} ) ) {
	        try {
		        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
	        }
	        catch {
		        Write-Warning $_.Exception.Message
		        EXIT
	        }
        }
        $Path = (Get-Location).Path
    }
    Process {
        If ( $Identity ) {
            ForEach ( $Database in $Identity ) {
                If ( $IncludePreExchange2010 ) {
                    Try {
                        Write-Verbose 'Obtaining Exchange 2010 and PreExchange 2010 Public Folder Databases'
                        $PFDatabase = Get-PublicFolderDatabase -Identity $Database -IncludePreExchange2010 -Status -ErrorAction STOP
                    }
                    Catch {
                        Write-Warning $_.Exception.Message
                    }
                }
                Else {
                    Try {
                        Write-Verbose 'Obtaining Exchange 2010 Public Folder Databases'
                        $PFDatabase = Get-PublicFolderDatabase -Identity $Database -Status -ErrorAction STOP
                    }
                    Catch {
                        Write-Warning $_.Exception.Message
                    }
                }
                ForEach ( $Database in $PFDatabase  ) {          
                    Write-Verbose "$($Database.Name) found on $($Database.Server.Name)"
                    Write-Verbose "Extracting data from $($Database.Name)"
                        $Data = Get-PublicFolderStatistics -Server $($Database.Server.Name) -ResultSize Unlimited | 
                        Where-Object {$_.DatabaseName -eq "$($Database.Name)"}            
                    Write-Verbose 'Data extraction completed'
                    Write-Verbose "Calculating FolderCount, Item Count and Deleted Item Count for $($Database.Name)"                                    
                        $Property = @{  "Server"                 = $Database.Server.Name
                                        "Folder Count"           = $Data.Count
                                        "Item Count"             = ($Data | Measure-Object -Property ItemCount -Sum).Sum
                                        "Deleted Item Count"     = ($Data | Measure-Object -Property DeletedItemCount -Sum).Sum
                                        "Public Folder Database" = $Database.Name
                                        "Database Size"          = $Database.DatabaseSize
                                        "White Space"            = $Database.AvailableNewMailboxSpace
                                    }                
                        $Object = New-Object -TypeName PSCustomObject -Property $Property            
                    Write-Verbose "Calculatio complete for $DatabaseName"
                    Write-Verbose 'Outputting Data'           
                        Write-Output $Object | Select "Server", "Public Folder Database", "Item Count", "Deleted Item Count", "Database Size", "White Space"        
                    Write-Verbose 'Output Successful'
                }   
            }
        }
        ElseIf ( $Server ) {
            ForEach ( $Server in $Server ) {         
                If ( $IncludePreExchange2010 ) {
                    Try {
                        Write-Verbose "'Obtaining Exchange 2010 and PreExchange 2010 Public Folder Databases from $Server"
                        $PFDatabase = Get-PublicFolderDatabase -Server $Server -IncludePreExchange2010 -Status -ErrorAction STOP
                    }
                    Catch {
                        Write-Warning $_.Exception.Message
                    }
                }
                Else {
                    Try {
                        Write-Verbose "Obtaining Exchange 2010 Public Folder Databases from $Server"
                        $PFDatabase = Get-PublicFolderDatabase -Server $Server -Status -ErrorAction STOP         
                    }
                    Catch {
                        Write-Warning $_.Exception.Message
                    }
                }
                ForEach ( $Database in $PFDatabase  ) {          
                    Write-Verbose "$($Database.Name) found on $($Database.Server.Name)"
                    Write-Verbose "Extracting data from $($Database.Name)"
                        $Data = Get-PublicFolderStatistics -Server $($Database.Server.Name) -ResultSize Unlimited | 
                        Where-Object {$_.DatabaseName -eq "$($Database.Name)"}            
                    Write-Verbose 'Data extraction completed'
                    Write-Verbose "Calculating FolderCount, Item Count and Deleted Item Count for $($Database.Name)"                                    
                        $Property = @{  "Server"                 = $Database.Server.Name
                                        "Folder Count"           = $Data.Count
                                        "Item Count"             = ($Data | Measure-Object -Property ItemCount -Sum).Sum
                                        "Deleted Item Count"     = ($Data | Measure-Object -Property DeletedItemCount -Sum).Sum
                                        "Public Folder Database" = $Database.Name
                                        "Database Size"          = $Database.DatabaseSize
                                        "White Space"            = $Database.AvailableNewMailboxSpace
                                    }                
                        $Object = New-Object -TypeName PSCustomObject -Property $Property            
                    Write-Verbose "Calculatio complete for $DatabaseName"
                    Write-Verbose 'Outputting Data'           
                        Write-Output $Object | Select "Server", "Public Folder Database", "Item Count", "Deleted Item Count", "Database Size", "White Space"        
                    Write-Verbose 'Output Successful'
                }               
            
            }
        }
    }
    End{
        Write-Verbose "Start End Block"
        Write-Verbose "Completed End Block"
    }
}
