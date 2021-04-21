$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
add-type -name win -member $t -namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)

while($true){
	$webrequest = (Invoke-WebRequest -URI "https://status.office.com" -UseBasicParsing -TimeoutSec 60)
	if($webrequest.RawContent -Match "Title:"){
		$wshell = New-Object -ComObject Wscript.Shell
		$wshell.Popup("Issues detected! Check status.office.com!!!",0,"Done",0x1)
	}
	sleep(3600)
}
