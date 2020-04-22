; insert comma 
; 1,000 units of numerical column
setComma(number) {
	numberString 	:= ""
	dot 		:= Strsplit(number , ".")	
	SplitNumber 	:= StrSplit(dot[1],"")
	maxcount 	:= SplitNumber.count()+1
	
	loop % SplitNumber.count()
		numberString :=  (mod(A_index,3) = 0 and maxcount > 2 and IsNumber(SplitNumber[maxcount-2]) ? "," : "" ) . SplitNumber[maxcount-=1]  .  numberString
		
	numberString := numberString . ( dot[2] ? "." dot[2] : "" )
	
	return numberString
}

; Make sure the numbers are correct
IsNumber(val) {
	if val is number
		return true 
	Else
		return false
}
