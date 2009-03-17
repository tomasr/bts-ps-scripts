
$script:bound = 2
$script:started = 4
$script:controlRecvLoc = 2
$script:controlInst = 2

function script:get-assemblyfilter([string]$assembly) {
   # The BizTalk WMI provider uses separate properties for each
   # part of the assembly name, so break it up to make it easier to handle
   $parts = $assembly.Split((',', '='))
   $filter = "AssemblyName='$($parts[0])'"
   if ( $parts.Count -gt 1 ) {
      for ( $i=1; $i -lt $parts.Count; $i += 2 ) {
         $filter = "$filter and Assembly$($parts[$i].trim())='$($parts[$i+1])'"
      }
   }
   return $filter
}
function script:find-orch([string]$name, [string]$assembly) {
   # We want to be able to find orchestrations by
   # name and/or assembly. That way we can control
   # all orchestrations in a single assembly in one call
   $filter = ""
   if ( ![String]::IsNullOrEmpty($name) ) {
      $filter = "Name='$name'"
      if ( ![String]::IsNullOrEmpty($assembly) ) {
         $filter = "$filter and $(get-assemblyfilter $assembly)"
      }
   } else {
      $filter = $(get-assemblyfilter $assembly)
   }
   get-wmiobject MSBTS_Orchestration `
      -namespace 'root\MicrosoftBizTalkServer' `
      -filter $filter
}

function start-orch([string]$name, [string]$assembly) {
   $orch = (find-orch $name $assembly)
   $orch | ?{ $_.OrchestrationStatus -eq $bound } | %{
      write-host "Enlisting $($_.Name)..."
      [void]$_.Enlist()
   }
   $orch | ?{ $_.OrchestrationStatus -ne $started } | %{
      write-host "Starting $($_.Name)..."
      [void]$_.Start($controlRecvLoc, $controlInst)
   }
}

function stop-orch([string]$name, [string]$assembly, [switch]$unenlist = $false) {
   $orch = (find-orch $name $assembly)
   $orch | ?{ $_.OrchestrationStatus -eq $started } | %{
      write-host "Stopping $($_.Name)..."
      [void]$_.Stop($controlRecvLoc, $controlInst)
      if ( $unenlist ) {
         [void]$_.Unenlist()
      }
   }
}

#start-orch -assembly 'Mimosa.BizTalk, Version=1.0.0.0, Culture=neutral, PublicKeyToken=50b7b2906e3f8aa5'
#stop-orch -assembly 'Mimosa.BizTalk, Version=1.0.0.0, Culture=neutral, PublicKeyToken=50b7b2906e3f8aa5' -unenlist
