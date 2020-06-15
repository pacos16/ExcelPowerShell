
#Excel options
$excel = New-Object -ComObject excel.application
$excel.application.Visible = $true
$excel.Displayalerts = $false
$book = $excel.workbooks.Add()
$sheet = $book.worksheets.Item(1)
$sheet.name = "Articulos"
$sheet.activate
$row=3
#Access Options
$path = "C:\nouges\ges2003.mdb"
$cn = New-Object -ComObject ADODB.Connection
$command = new-object -ComObject ADODB.Command
$cn.Open("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = $path")
$command.activeConnection = $cn
$command.commandText ='select * from articulos where ref like :ref'
$command.commandType = 1
$paramRef = $command.createParameter('ref',3,1,4)
$command.parameters.append($paramRef)

#Code
Get-ChildItem "c:/nouges/img" | ForEach-Object{
    $ref= $_.Name.Substring(0,3)
    $paramRef.value = $ref
    $resultset=$command.execute() 
    try{
        $resultset.moveFirst();

        $hashMap= @{
            'image'= $_.Name
            'descripcion'= $resultset.fields.item("desc").value
            'ref'= $resultset.fields.item("ref").value
        }
        $object = New-object -TypeName PSObject -property $hashMap;
        Write-Output $object
        $sheet.Cells.item($row,1) = $object.ref
        $sheet.Cells.item($row,2) = $object.image
        $sheet.Cells.item($row,3) = $object.descripcion
        $row++
    }catch{

    }
}

$cn.close();
