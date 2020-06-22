if ($env:Processor_Architecture -ne "x86") { 
&"$env:windir\syswow64\windowspowershell\v1.0\powershell.exe" -noninteractive -noprofile -file $myinvocation.Mycommand.path -executionpolicy bypass
exit
}

$path = "C:\nouges\ges2003.mdb"
$adOpenForwardOnly= 0
$adLockOptimistic = 3

$cn = New-Object -ComObject ADODB.Connection
$rs = New-Object -Comobject ADODB.Recordset

$cn.Open("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = $path")
$rs.Open("select * from articulos" , $cn, $adOpenForwardOnly, $adLockOptimistic)

$rs.moveFirst()

$array=@()

do{
    $hashMap = @{
        'descripcion' =$rs.Fields.Item(4).value;
        'ref' = $rs.Fields.Item(3).value
    } ;
    $object = New-object -TypeName PSObject -property $hashMap;
    $array += $object;
    $rs.movenext();
} until ($rs.EOF -eq $True)

$rs.close();
$cn.close();

$json= ConvertTo-json $array

$json > "c://nouges/articulos.json"
