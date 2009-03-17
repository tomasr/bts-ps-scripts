#
# declare our parameters: the action to take
#
param(
   [string] $action=$(throw 'need action'),
   [string] $name,
   [string] $status
)


#
# Helper function to list all WMI objects of
# a given type
#
function bts-list-objs($kind)
{
   get-wmiobject $kind `
      -namespace 'root\MicrosoftBizTalkServer'
}

#
# display all information about a receive port
#
function bts-display-recv-port($port)
{
   $portname = $port.Name
   $portname
   $recvlocs = get-wmiobject MSBTS_ReceiveLocation `
      -namespace 'root\MicrosoftBizTalkServer' `
      -filter "ReceivePortName='$portName'"

   $recvlocs | ft Name, AdapterName, InboundTransportURL, IsDisabled
}

#
# display all information about a send port group
#
function bts-display-send-port-group($group)
{
   $groupname = $group.Name
   
   $ports = get-wmiobject MSBTS_SendPortGroup2SendPort `
      -namespace 'root\MicrosoftBizTalkServer' `
      -filter "SendPortGroupName='$groupName'"

   $groupname
   foreach ($port in $ports) { "`t" + $port.SendPortName }
}

#
# enable or disable the specified
# receive location
#
function bts-set-recv-loc-status($name, $status)
{
   $recvloc = get-wmiobject MSBTS_ReceiveLocation `
      -namespace 'root\MicrosoftBizTalkServer' `
      -filter "Name='$name'"
   
   switch ( $status )
   {
      'disable' { [void]$recvloc.Disable() }
      'enable' { [void]$recvloc.Enable() }
   }
}

#
# controls a send port
#
function bts-set-send-port-status($name, $status)
{
   $sendport = get-wmiobject MSBTS_sendPort `
      -namespace 'root\MicrosoftBizTalkServer' `
      -filter "Name='$name'"

   switch ( $status )
   {
      'start' { [void]$sendport.Start() }
      'stop' { [void]$sendport.Stop() }
      'enlist' { [void]$sendport.Enlist() }
      'unenlist' { [void]$sendport.UnEnlist() }
   }
}

function is-valid-opt($check, $options)
{
   foreach ( $val in $options )
   {
      if ( $check -eq $val ) { 
         return $true
      }
   }
}

#
# main script
#
switch ( $action )
{
   'list-send-ports' {
      bts-list-objs('MSBTS_SendPort') | ft Name, PTAddress, Status
   }
   'list-send-port-groups' {
      bts-list-objs('MSBTS_SendPortGroup') | %{ bts-display-send-port-group($_) }
   }
   'list-recv-ports' {
      bts-list-objs('MSBTS_ReceivePort') | %{ bts-display-recv-port($_) }
   }
   'set-recv-loc-status' {
      if ( !(is-valid-opt $status ('enable', 'disable')) )
      {
         throw 'invalid status'
      }
      if ( ($name -eq '') -or ($name -eq $null) )
      {
         throw 'you must supply the receive location name'
      }
      bts-set-recv-loc-status $name $status
   }
   'set-send-port-status' {
      if ( !(is-valid-opt $status ('start', 'stop', 'enlist', 'unenlist')) )
      {
         throw 'invalid status'
      }
      if ( ($name -eq '') -or ($name -eq $null) )
      {
         throw 'you must supply the send port name'
      }
      bts-set-send-port-status $name $status
   }
}

