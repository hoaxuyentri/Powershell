###Define input data
$FuelType = 'DO 0,05S-II'
$server = "LAPTOP310\SQLLUANTT"
$database = "MIN_NPM"

###Get website html content
$WebResponseObj = Invoke-WebRequest -Uri "https://www.petrolimex.com.vn" -UseBasicParsing
$Content = $WebResponseObj.Content

###Define input parameters for store procedure
$String = $Content.Substring($Content.IndexOf($FuelType), $FuelType.Length + 38)
$Zone1 = ($String.Substring($FuelType.Length + 15, 6)).Replace(".","")
$Zone2 = ($String.Substring($String.Length - 6, 6)).Replace(".","")
#Write-Output $Zone1
#Write-Output $Zone2

$MFirstDate = (Get-date).AddDays(1-(Get-date).Day).ToString("yyyy-MM-dd")
$MLastDate = (((Get-date).AddDays(1-(Get-date).Day)).AddMonths(1)).AddDays(-1).ToString("yyyy-MM-dd")
#Write-Output $MFirstDate
#Write-Output $MLastDate

###Connect to server and execute store procedure
$scon = New-Object System.Data.SqlClient.SqlConnection
$scon.ConnectionString = "Data Source=$server;Initial Catalog=$database;Integrated Security=true"
        
$cmd = New-Object System.Data.SqlClient.SqlCommand
$cmd.Connection = $scon
$cmd.CommandTimeout = 0

### Tie parameters using SqlParameters
$cmd.CommandText = "DECLARE @ID INT EXEC PP_UI_UPDATE_FUELRATE 1,'$MFirstDate','$MLastDate',$Zone1,$Zone2,@ID"

try
 {
   $scon.Open()
   $cmd.ExecuteNonQuery() | Out-Null
 }
catch [Exception]
 {
   Write-Warning $_.Exception.Message
 }
finally
 {
   $scon.Dispose()
   $cmd.Dispose()
 }
