On Error Resume Next

set WshShell = WScript.CreateObject("WScript.Shell")

'
' Obtain currently logged on user's DN
' stored in the value: ADSysInfo.UserName
'
Set ADSysInfo = CreateObject("ADSystemInfo")

'
' Query user's SIP URI and home server DN
' using the following WMI class:
'
'	class MSFT_SIPESUserSetting
'	{
'		string DisplayName;
'		boolean Enabled;
'		string HomeServerDN;
'		[key] string InstanceID;
'		string PrimaryURI;
'		string TargetServerDNIfMoving;
'		string UserDN;
'	};
'
Set oUsers = GetObject("winmgmts:\\").ExecQuery("SELECT * FROM MSFT_SIPESUserSetting WHERE UserDN = '" & ADSysInfo.UserName & "'")
' Since the user DN is unique, we're expecting a single result.
for each User in oUsers
	sipURI = User.PrimaryURI
	poolDN = User.HomeServerDN
next

' Exit if user is not enabled for Live Communications.
if IsEmpty(sipURI) then 
	Wscript.echo "ERROR: user is not enabled for Live Communications"
	Wscript.Quit
end if

' Exit if user is orphaned.
if IsNull(poolDN) then 
	Wscript.echo "ERROR: user is not assigned to a home server"
	Wscript.Quit
end if

'
' Convert home server DN into FQDN
' using the following WMI class:
'
'	class MSFT_SIPPoolSetting
'	{
'  		string BackEndDBPath;
'  		[key] string InstanceID;
'  		uint32 MajorVersion;
'  		uint32 MinorVersion;
'  		string PoolDisplayName;
'  		string PoolDN;
'  		string PoolFQDN;
'  		string [] PoolMemberList;
'  		string PoolType;
'	};
'
Set oServers = GetObject("winmgmts:\\").ExecQuery("SELECT * FROM MSFT_SIPPoolSetting WHERE PoolDN = '" & poolDN & "'")
' Since the pool DN is unique, we're only expecting a single FQDN.
for each HS in oServers
	poolFQDN = HS.PoolFQDN
next

' Exit if user's home server FQDN can not be determined.
if IsEmpty(poolFQDN) then 
	Wscript.echo "ERROR: failed to determine user's home server FQDN"
	Wscript.Quit
end if

'
' Retrieve user's domain
'
domain = WshShell.ExpandEnvironmentStrings("%USERDOMAIN%")

'
' Parse the SIP URI into the format:
' 'sip:user@company.com' -> 'user@company.com %domain%\user'
' where %domain% is the user's domain name
'
Set myRegExp = New RegExp
myRegExp.IgnoreCase = True
myRegExp.Global = True
myRegExp.Pattern = "sip:([^~]+)@([^~]+)"
replaceString = "$1@$2 " & domain & "\$1"
userURI = myRegExp.Replace(sipURI, replaceString)

'
' Configure Communicator registry keys
'
WshShell.RegWrite "HKCU\Software\Microsoft\Communicator\UserMicrosoft RTC Instant Messaging", userURI, "REG_SZ"
WshShell.RegWrite "HKCU\Software\Microsoft\Communicator\ServerAddress", poolFQDN, "REG_SZ"
WshShell.RegWrite "HKCU\Software\Microsoft\Communicator\ConfigurationMode", 1, "REG_DWORD"
' Set type to 2 for TCP or set type to 4 for TLS
const TransportType = 2
WshShell.RegWrite "HKCU\Software\Microsoft\Communicator\Transport", TransportType, "REG_DWORD"

'
' Start Communicator.exe so user doesn't experience any delays
'
WshShell.Run "Communicator.exe", 1
