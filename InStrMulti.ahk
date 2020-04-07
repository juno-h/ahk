 ; 배열 Needle_array 에서 Haystack 의 값 유무 반환
 
InStrMulti(Haystack , Needle_array)
{
	loop % Needle_array.count()
	{
		matched := instr(Haystack, Needle_array[A_index]) ? true : false
		if (matched = true)
		break
	}
	return matched
}
