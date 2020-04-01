; insert comma 
setComma(number) {
	temp := ""
	SplitNumber := StrSplit(number,"") , maxcount := SplitNumber.count()+1
	loop % SplitNumber.count()
		temp :=  (mod(A_index,3) = 0 and maxcount > 2 and IsNumber(SplitNumber[maxcount-2]) ? "," : "" ) . SplitNumber[maxcount-=1]  .  temp
	return temp
}
