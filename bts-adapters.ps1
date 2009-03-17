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
# get objects related to an adapter
# 
function bts-get-related($adapterName, $kind)
{
   get-wmiobject $kind `
      -namespace 'root\MicrosoftBizTalkServer' `
      -filter "AdapterName='$adapterName'"
}

#
# dump adapter information to the console
#
function bts-show-adapter($adapter)
{
   '--- ' + $adapter.Name + ' ---'
   $sendHandlers = @(bts-get-related $adapter.Name MSBTS_SendHandler2)
   if ( $sendHandlers.Length -gt 0 )
   {
      'Send Handlers:'
      for ( $i =0; $i -lt $sendHandlers.Length; $i++ )
      {
         '   ' + $sendHandlers[$i].HostName
      }
   }
   
   $recvHandlers = @(bts-get-related $adapter.Name MSBTS_ReceiveHandler)
   if ( $recvHandlers.Length -gt 0 )
   {
      'Receive Handlers:'
      for ( $i =0; $i -lt $recvHandlers.Length; $i++ )
      {
         '   ' + $recvHandlers[$i].HostName
      }
   }
   ''
}

#
# main script
#
bts-list-objs MSBTS_AdapterSetting | %{ bts-show-adapter $_ }


