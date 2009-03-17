
$hosts = get-wmiobject MSBTS_HostInstance -namespace 'root\MicrosoftBizTalkServer'
$hosts | sort HostName | ft HostName, HostType