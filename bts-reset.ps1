
#
# declare two switch parameters: -start and -stop
#
param([switch] $start, [switch] $stop)

#
# get list of application hosts
#
function get-apphosts
{
	get-wmiobject MSBTS_HostInstance `
		-namespace 'root\MicrosoftBizTalkServer' `
		-filter HostType=1
}

#
# stop the given host
#
function stop-host($apphost)
{
	$hostname = $apphost.HostName
	if ( $apphost.ServiceState -ne 1 )
	{
		"Stopping Host $hostname ..."
		[void]$apphost.Stop()
	}
}

#
# start the given host
#
function start-host($apphost)
{
	$hostname = $apphost.HostName
	if ( $apphost.ServiceState -eq 1 )
	{ 
		"Starting Host $hostname ..."
		[void]$apphost.Start()
	}
}

#
# main script
#

if ( !($stop) -and !($start) )
{
	$stop = $true
	$start = $true
}

if ( $stop )
{
	get-apphosts | %{ stop-host($_) }
}

if ( $start )
{
   get-apphosts | %{ start-host($_) }
}
