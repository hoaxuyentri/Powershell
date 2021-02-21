$Hide = 0
$Normal = 1
$Minimized = 2
$Maximized = 3
$ShowNoActivateRecentPosition = 4
$Show = 5
$MinimizeActivateNext = 6
$MinimizeNoActivate = 7
$ShowNoActivate = 8
$Restore = 9
$ShowDefault = 10
$ForceMinimize = 11

#Specify an interwebs address
$URL="https://www.petrolimex.com.vn/"

#Create internetexplorer.application object
$IE=new-object -com "InternetExplorer.Application"

#Set some parameters for the internetexplorer.application object
$IE.TheaterMode = $False
$IE.AddressBar = $True
$IE.StatusBar = $False
$IE.MenuBar = $True
$IE.FullScreen = $False
$IE.visible = $True

#Navigate to the URL
$IE.navigate2($URL,$null,$true)

#the C#-style signature of an API function
$code = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'

#add signature as new type to PowerShell (for this session)
$type = Add-Type -MemberDefinition $code -Name myAPI -PassThru

#Magic:
$type::ShowWindowAsync($IE.HWND[0], $Maximized) | Out-Null

do {sleep 5} until (-not ($IE.Busy))

# Take A ScreenShot
[Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
function screenshot([Drawing.Rectangle]$bounds, $path) {
	$bmp = New-Object Drawing.Bitmap $bounds.width, $bounds.height
	$graphics = [Drawing.Graphics]::FromImage($bmp)

	$graphics.CopyFromScreen($bounds.Location, [Drawing.Point]::Empty, $bounds.size)

	$bmp.Save($path)

	$graphics.Dispose()
	$bmp.Dispose()
}

$path = "C:\Luan data\Fuel price " + (Get-date).ToString("yyyy-MM-dd HHmmss") + ".jpg"

$bounds = [Drawing.Rectangle]::FromLTRB(0, 0, 1360, 770)
screenshot $bounds $path

$shellapp = New-Object -ComObject "Shell.Application"
$ShellWindows = $shellapp.Windows()
for ($i = 0; $i -lt $ShellWindows.Count; $i++)
{
 if ($ShellWindows.Item($i).FullName -like "*iexplore.exe")
  {
  	$ie = $ShellWindows.Item($i)
  	$ie.quit()
  	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ie) | Out-Null
  	[System.GC]::Collect()
  	[System.GC]::WaitForPendingFinalizers()
  }
}