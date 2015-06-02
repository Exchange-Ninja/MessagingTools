# MessagingTools

This PowerShell Module uses the Microsoft.Exchange.Management.PowerShell.E2010 PSSnapin. It is built and designed to work specifically with Exchange 2010 SP3 RU5+

It Requires PowerShell V2.0 and Exchange Management Tools if not ran on an Exchang Server directly. 

The module contains the following functions.

Get-MessageCount
Get-PublicFolderReport

Every funvtion requires an active PSSession with an Exchange Server. The functions have built in validations that Microsoft.Exchange.Management.PowerShell.E2010 PSSnapin is available. 
