add-type -AssemblyName System.Data.OracleClient

$username = "##"
$password = "####"
$data_source = "####"
$connection_string = "User Id=$username;Password=$password;Data Source=$data_source"
$orders = Get-Content -path 'C:\Users\##\Desktop\orders.txt'
$lastline = $orders[-1]

$source = for ($i = 0; $i -lt $orders.Count; $i++) {
  $number = $orders[$i]
  if($number -eq $lastline){
     $number = "'$number'"
  } else { 
     $number = "'$number',"
  }
  $number
}

$currentDate = (Get-Date).ToString("dd-MMM-yy").ToUpper()
$currentDate2 = "'%$currentDate%'"

$statement = "select workorderid from workorder w join product p on w.productid = p.productid where w.creationdate like $currentDate2"

try {
    $con = New-Object System.Data.OracleClient.OracleConnection($connection_string)
    $con.Open()

    $cmd = $con.CreateCommand()
    $cmd.CommandText = $statement

    $result = $cmd.ExecuteReader()
    $ordersfinal = @()

    while ($result.Read()) {
        $workorderid = $result["workorderid"]
        $hexValue = [int]$workorderid
        $ordersfinal += $hexValue.ToString("X")
    }

    $result.Close()
    $ordersfinal | Out-File "C:\Users\knyka\Desktop\table.txt"
} catch {
    Write-Error ("Database Exception: {0}`n{1}" -f $con.ConnectionString, $_.Exception.ToString())
} finally {
    if ($con.State -eq 'Open') { $con.Close() }
}