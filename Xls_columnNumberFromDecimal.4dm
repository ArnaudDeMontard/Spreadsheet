//Xls_columnNumberFromDecimal (numColonneDecimal_l) -> txt
//$0 : $1 converti en notation colonne tableur
//  decimal * 1, 2...26, 27, 28...703
//  tableur * A, B... Z, AA, AB...AAA
//méthode ré-entrante

//µ arnaud * 25/11/06 * adapté, étendu à plus de 256 colonnes
//© Laurent Jolly * 24/11/06

var $0 : Text
var $1 : Integer

var $out_t : Text
var $col_l : Integer

var $unite : Integer
var $base : Integer

If (False)  //tests unitaires
	SET ASSERT ENABLED(True)
	ASSERT(Xls_columnNumberFromDecimal(9)="I")
	ASSERT(Xls_columnNumberFromDecimal(27)="AA")
	ASSERT(Xls_columnNumberFromDecimal(122)="DR")
	ASSERT(Xls_columnNumberFromDecimal(258)="IX")
	ASSERT(Xls_columnNumberFromDecimal(677)="ZA")
	ASSERT(Xls_columnNumberFromDecimal(703)="AAA")
	ASSERT(Xls_columnNumberFromDecimal(1024)="AMJ")
	ASSERT(Xls_columnNumberFromDecimal(2040)="BZL")
	ASSERT(Xls_columnNumberFromDecimal(2586)="CUL")  // ;-)
End if 

$out_t:=""
ASSERT(Count parameters>0)
$col_l:=$1
ASSERT($col_l>0)
$unite:=(($col_l-1)%26)+1
$base:=(($col_l-1)\26)
//$unite:=Mod($col_l-1; 26)+1
//$base:=Int(($col_l-1)/26)
If ($base>0)
	$out_t:=Xls_columnNumberFromDecimal($base)+$out_t
End if 
$0:=$out_t+Char($unite+64)
