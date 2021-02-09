$path = "C:\nouges\ges2003.mdb"
$adOpenForwardOnly= 0
$adLockOptimistic = 3

$path = "C:\nouges\temp2017.mdb"
$cn = New-Object -ComObject ADODB.Connection
$command = new-object -ComObject ADODB.Command
$cn.Open("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = $path")
$command.activeConnection = $cn
$command.commandText ='select * from articulos where ref like :ref'
$command.commandType = 1
$paramRef = $command.createParameter('ref',129,1,4)
$command.parameters.append($paramRef)

$array=@()
Get-ChildItem "C:\\nouges\jpg_2020\*.jpg" | ForEach-Object{
    $paramRef.value = $_.name.Split('.')[0]
    $rs=$command.execute() 

    $rs.moveFirst()
    $hashMap = @{
        'medidas' =$rs.Fields.Item(6).value;
        'descripcionExp' =$rs.Fields.Item(5).value;
        'descripcion' =$rs.Fields.Item(4).value;
        'ref' = $rs.Fields.Item(3).value
        'nombre' =$rs.Fields.Item(31).value;
        'categoria' =$rs.Fields.Item(32).value;
        'colores' =$rs.Fields.Item(33).value;
        'pvp' = [math]::Round([float] $rs.Fields.Item(21).value,2)
    } ;
    $object = New-object -TypeName PSObject -property $hashMap;
    if($object.ref -ne $null){
        $array += $object;
    }
    
}

$arrayCSV=@()
$csv = Import-Csv "C:\nouges\Scripts\importador.csv" -Encoding UTF8
$simple = $csv[0]
$compuesto = $csv[1]
$compuesto.sku = ""
$compuesto.'Clase de envío' = ""
$simple.'Clase de envío' = ""
$simple.categorías = ""
$compuesto.categorías = ""
$compuesto.'Precio normal'= 0
$simple.'Precio normal'= 0
$simple.'Valor(es) del atributo 1'= ""
$Compuesto.'Valor(es) del atributo 2' = ""
$Simple.'Valor(es) del atributo 2' = ""



$array | ForEach-Object{
    $sizeSplit=$null
    if($null -ne $_.medidas){
        $sizeSplit=$_.medidas.split('-')
    }else{
        $sizeSplit=$null
    }

    $colorSplit=$null
    if($null -ne $_.colores){
        $colorSplit=$_.colores.split(',')
    }else{
        $colorSplit=$null
    }
    
    
    $articulo = $_
    $productCsvAux = @()
    $productCsvAux2 = @()

    $simple.id = [String] $articulo.ref 
    $simple.Imágenes = "https://sombrerossiver.es/wp-content/uploads/"+$articulo.ref+".jpg"
    $simple.nombre= $articulo.nombre
    
    $simple.'Descripción corta' = $articulo.descipcion
    $simple.Descripción = $articulo.descripcionExp
    $simple.'Clase de impuesto' =""
    $simple.categorías = $articulo.categoria
    $compuesto.categorías = ""
    if($null -ne $sizeSplit[1]){
        
        for($i=1; $i -le $sizeSplit.Length; $i++){
       
            
            $compuesto.id = $articulo.ref + "-" + $i
            $compuesto.tipo = 'variation'
            $compuesto.nombre= $articulo.nombre
            $compuesto.'Descripción corta' = ""
            $compuesto.Descripción = ""
            $compuesto.'Precio normal'= $articulo.pvp
            $compuesto.'Clase de envío' = ""
            $compuesto.Imágenes = ""
            $compuesto.'Clase de impuesto'= "parent"
            $compuesto.Posición = $i 
            $Compuesto.'Valor(es) del atributo 1' = $sizeSplit[$i-1]
            $compuesto.'Atributo visible 1' = ""
            $Compuesto.'Valor(es) del atributo 2' = $articulo.Colores

            $productCsvAux+= $compuesto.psobject.copy()
            
        }
        $simple.Tipo = 'variable'
        $simple.'Nombre del atributo 1' = 'Talla' 
        $simple.'Atributo global 1' = 1
        $simple.'Valor(es) del atributo 1' = $articulo.medidas
        $simple.'Atributo visible 1' = 1
        $simple.'Precio normal'= ""
        $simple.categorías = $articulo.categoria
        $contador = 1;
        if($null -ne $colorSplit[1]){
            $colorSplit | ForEach-Object{
                $color = $_
                $productCsvAux | ForEach-Object{
                    $_.'Valor(es) del atributo 2' = $color
                    $_.id=$articulo.ref + "-" + $contador
                    $_.'Nombre del atributo 2' = 'Color'
                    $_.'Atributo global 2' = 1
                    $_.posición = $contador
                    $contador++
                    $productCsvAux2 += $_.psobject.copy()
                }
            }
            $productCsvAux=$productCsvAux2
        }
        $simple.'Nombre del atributo 2' = 'Colores' 
        $simple.'Atributo global 2' = 1
        $simple.'Valor(es) del atributo 2' = $articulo.colores
        $simple.'Atributo visible 2' = 1
        

    }else{
        $simple.'Precio normal'= $articulo.pvp
        $simple.Tipo= 'simple'
        $simple.'Atributo visible 1' = ""
        $simple.'Valor(es) del atributo 1'= $articulo.medidas
        $simple.'Atributo global 1' = ""
        $simple.categorías = $articulo.categoria
        $simple.'Valor(es) del atributo 2'= $articulo.Colores
        if($null -ne $colorSplit[1]){
            $contador=1;
            $colorSplit | ForEach-Object{

                $compuesto.id = $articulo.ref + "-" + $contador
                $compuesto.tipo = 'variation'
                $compuesto.nombre= $articulo.nombre
                $compuesto.'Descripción corta' = ""
                $compuesto.Descripción = ""
                $compuesto.'Precio normal'= $articulo.pvp
                $compuesto.'Clase de envío' = ""
                $compuesto.Imágenes = ""
                $compuesto.'Clase de impuesto'= "parent"
                $compuesto.Posición = $contador 
                $Compuesto.'Valor(es) del atributo 1' = $articulo.medidas
                $Compuesto.'Valor(es) del atributo 2' = $_
                $compuesto.'Atributo visible 2' = ""
                $contador++
                $productCsvAux+= $compuesto.psobject.copy()
            }

            $simple.Tipo= 'variable'
        }
        
    }

    $arrayCSV += $simple.psobject.copy()
    $productCsvAux | ForEach-Object{
        $arrayCsv += $_.psobject.copy()
    }

}
$arrayCsv |Export-Csv .\importadorSiver.csv -Encoding UTF8
$rs.close();
$cn.close();

