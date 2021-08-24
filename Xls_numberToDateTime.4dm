//Xls_numberToDateTime (obj) -> obj
//converts a spreadsheet number to 4D date and time
//$0.xxx :
//  date : a date
//  time : a time
//$1.xxx :
//  XLnumber : number to convert (decimals = seconds)
//  origin : optional origin date, default is 1900-01-01
//    (mac is 1904-01-01, windows is 1900-01-01)

var $0 : Object
var $1 : Object

var $out_o : Object
var $in_o : Object

$out_o:=New object
$in_o:=$1
If ($in_o.XLnumber#Null)
	$origin_d:=!1900-01-01!  //PC
	If ($in_o.origin#Null)
		$origin_d:=$in_o.origin
	End if 
	$days_l:=Int($in_o.XLnumber)
	$seconds_l:=(86400*(Dec($in_o.XLnumber)))
	$out_o.date:=$origin_d+$days_l
	$out_o.time:=Time(Time string($seconds_l))
End if 
$0:=$out_o
