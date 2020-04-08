; Remove Duplicate value
RemoveDuplicateArray(array){
	temp := []
	for index , val in array {
		check := false
		for idx , tempval in temp {
			if (tempval = val) {
				check := true
				break
			}
		}
		(val and !check) ? temp.push(val) : ""
	}
	return temp
}
