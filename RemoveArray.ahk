; parameter로 받은 value를 array 함수내 존재시 제거

RemoveArray(byref array , value)
{
	for key , val in array
		(val = value) ? array.delete(key) : ""
	return array
}
