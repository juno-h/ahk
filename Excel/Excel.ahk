Class excel
{
	__new(mode)
	{
		try {
			this.xl 	:=  (mode == "active") ? ComObjActive("Excel.Application")
						: 	(mode == "create") ? ComObjCreate("Excel.Application") 
						: 	""
		} catch e {
			msgbox , 16 , error, please open excel ..
			Return 
		}
				
	}
	Selectioncopy()
	{	
		this.checkdata()
		try {
			this.xl.Selection.Copy
			this.data 		:= Clipboard
			Clipboard 		:= ""
		} catch e{
			msgbox , 16 , error, please select area ..
			Return
		}
		return this.data
	}
	checkdata()
	{
		AtvCell := this.xl.ActiveCell.Address	
		if (this.xl.Range(AtvCell).value = "")
		{	
			msgbox, 선택셀의 데이터가 유효하지 않습니다.`n다시 선택해주세요
			return
		}
		return this
	}
	array()
	{
		array 		 	:= []
		array_stay	 	:= []
		field_lastdata	 := ""
		Columnscount := this.xl.Selection.Columns.Count
		this.data := StrReplace(this.data ,"""","")
		this.data := StrReplace(this.data ,"'","")	
		loop , parse, % this.data ,`n ,`r
		{
			if !Strlen(A_loopfield) || (A_loopfield = "")
				continue
			field := StrSplit(A_loopfield,"`t")
			if (field.count() = Columnscount) {
				array.push(field)
			} else {
				loop % field.count()
				{
					if (A_index = field.count() and ((field.count()+array_stay_count) <> Columnscount) )
					{
						field_lastdata := field[A_index]
					} else {
						(A_index = 1 and field_lastdata <> "" ) ? (field[A_index] := field_lastdata . " " . field[A_index]) : ""
						array_stay.push(field[A_index])
					}	
					(array_stay.count() = Columnscount) ? (array.push(array_stay) , field_lastdata := "" , array_stay := [] ) : ""
				}
				array_stay_count := array_stay.count()
			}			
		}		
		this.array := array
		return this
	}
	getData()
	{
		array 		 	:= []
		Columnscount 	:= this.getColumnCount()
		Rowcount 		:= this.getRowCount()
		this.data 		:= StrReplace(this.data ,"""","")
		this.data 		:= StrReplace(this.data ,"'","")	
		DataRow 		:= StrSplit(this.data,"`n")
		(DataRow[DataRow.count()] == "") ? DataRow.RemoveAt(DataRow.count()) : "" 
		if (Rowcount <> DataRow.count())
		{
			msgbox , 16, error , 전체 열과 DATA 열이 맞지 않습니다.`n 'enter' 가 삽입 되어 있는 지 확인하세요.`n%Rowcount% 
			exit
		}
		loop % DataRow.count()
		{
			field := StrSplit(DataRow[A_index],"`t")
			array.push(field)
		}		
		return array
	}

	getRowCount(mode := "")
	{
		if(this.xl == "")
			return -1
		else if (mode == "a")
			return this.xl.ActiveSheet.usedRange.rows.count								
		else if (mode == "o")
			return this.xl.workbook.sheets(this.current_sheet).usedRange.rows.count
		else 
			return this.xl.Selection.rows.Count
	}	

	getColumnCount(mode := "")
	{
		if(this.xl == "")
			return -1
		else if (mode == "a")
			return this.xl.ActiveSheet.usedRange.columns.count
		else if (mode == "o")
			return this.xl.workbook.sheets(this.current_sheet).usedRange.columns.count
		else 
			return this.xl.Selection.columns.Count
	}

	getAtiveWorkbookName()
	{
		if(this.xl == "")
			return -1
		else 
			return this.xl.ActiveWorkbook.Name
	}

	setSheetsCount()
	{
		return this.xl.Sheets.Count
	}
	
	setSheets()
	{
		return this.xl.Sheets.Add()
	}	
	setPaste(cell := "A1")
	{
		this.xl.Range(cell).pasteSpecial paste
		return 
	}
}

Class DataControl
{

	__new(data){
		this.data := data
		return this
	}

	Arraypos(){
		ArrayPos 		:= object()
		rowdata 		:= this.data
		loop , parse , rowdata ,`n , `r 
		{
			loopData := StrSplit(A_loopfield , "`t")
			ArrayPos["x" 		, A_index] := loopData[1]
			ArrayPos["y" 		, A_index] := loopData[2]
			ArrayPos["angle"	, A_index] := loopData[3]
			ArrayPos["stap" 	, A_index] := loopData[4]
			ArrayPos["ch" 		, A_index] := loopData[5]
		}
		this.Array := ArrayPos
		return ArrayPos
	}

	getChannel(mode := ""){

		ch 		:= object()
		rowdata := this.data

		if (mode = "")
		{
			loop , parse , rowdata ,`n , `r
			{
				loopData := StrSplit(A_loopfield , "`t")
				loopData[5] <> "" && !hasval( ch , loopData[5] ) ? ch.push(loopData[5]) : ""
			}
		} else if (mode = "IsNumber")
		{
			loop , parse , rowdata ,`n , `r
			{
				loopData := StrSplit(A_loopfield , "`t")
				var_number := loopData[5]
				if var_number is number
					hasval( ch , loopData[5] ) ? "" : ch.push(loopData[5])
			}
		} else if (mode = "IsNotNumber")
		{
			loop , parse , rowdata ,`n , `r
			{
				loopData := StrSplit(A_loopfield , "`t")
				var_number := loopData[5]
				if var_number
					if var_number is not number
						hasval( ch , loopData[5] ) ? "" : ch.push(loopData[5])
			}
		} else {
			loop , parse , rowdata ,`n , `r
			{
				loopData := StrSplit(A_loopfield , "`t")
				InStr(loopData[5] , mode) && !hasval( ch , loopData[5] ) ? ch.push(loopData[5]) : ""
			}
		}
		return 	ch
	}

	setArrayChannel(mode := ""){

		Channel := object()
		rowdata := this.data

		if (mode = "")
		{
			loop , parse , rowdata ,`n , `r
			{
				loopData := StrSplit(A_loopfield , "`t")
				isObject(Channel[loopData[5]]) ? "" : Channel[loopData[5]] := []
				Channel[loopData[5]].push({x : loopData[1] , y : loopData[2] , angle : loopData[3] , stap : loopData[4] })
			}
		} 
		return 	Channel
	}

	setArrayAngle(mode := ""){

		angle 	:= object()
		rowdata := this.data

		if (mode = "")
		{
			loop , parse , rowdata ,`n , `r
			{
				loopData := StrSplit(A_loopfield , "`t")
				isObject(angle[loopData[3]]) ? "" : angle[loopData[3]] := []
				angle[loopData[3]].push({x : loopData[1] , y : loopData[2] , stap : loopData[4] , ch : loopData[5] })
			}
		} 
		return 	angle
	}

	setArrayPos(mode){
		
		POS 	:= object()
		rowdata := this.data

		if (mode = "x")
		{
			loop , parse , rowdata ,`n , `r
			{
				loopData := StrSplit(A_loopfield , "`t"), point := loopData[1]
				point += 0 
				hasval( POS , point ) ? "" : POS.push(point)
				
			}
		} else if (mode = "y")
		{
			loop , parse , rowdata ,`n , `r
			{
				loopData := StrSplit(A_loopfield , "`t"), point := loopData[2]
				point += 0 
				hasval( POS , point ) ? "" : POS.push(point)
			}
		} 
		return 	POS
	}

	setArrayXYPos(){			; return array[ xpos "," ypos]
		XYPOS 	:= object()
		rowdata := this.data
		loop , parse , rowdata ,`n , `r
		{
			loopData := StrSplit(A_loopfield , "`t")
			isObject( XYPOS[ loopData[1] "," loopData[2] ] ) ? "" : XYPOS[ loopData[1] "," loopData[2] ] := []
			XYPOS[ loopData[1] "," loopData[2] ]	:=	{angle : loopData[3] , stap : loopData[4] , ch : loopData[5] }
		}
		return 	XYPOS
	}

	getPosGap(){
		row 	:= this.Arraypos()
		POS 	:= object()
		loop % row.x.count()
		{
			a := b:= A_index
			growth_x := 0
			growth_y := 0

			loop % row.x.count() + 1 - a
			{
				POS[ "x" , "front"	, row.x[a] "," row.y[a] , A_index ] := row.x[b+1] - row.x[b]
				POS[ "x" , "back" 	, row.x[a] "," row.y[a] , A_index ] := row.x[b] - row.x[b-1]
				POS[ "y" , "front" 	, row.x[a] "," row.y[a] , A_index ] := row.y[b+1] - row.y[b]
				POS[ "y" , "back" 	, row.x[a] "," row.y[a] , A_index ] := row.y[b] - row.y[b-1]

				; Gap 증가

				POS[ "x" , "growth"	, row.x[a] "," row.y[a] , A_index ] := growth_x := growth_x + (row.x[b+1] - row.x[b])
				POS[ "y" , "growth" , row.x[a] "," row.y[a] , A_index ] := growth_y := growth_y + (row.y[b+1] - row.y[b])
				b++
			}
		}
		return 	POS
	}

	setStapLength(){

		return	
	}

}


; GetTestData()
; {
; data = 
; (
; -2492.33	1285.73	180	1	PVIN8O
; -1971.33	1190.98	180	4	PVIN8O
; -936.23	1187.33	180	7	SGND
; -2036.53	1186.73	180	3	PVIN8
; -2427.33	1112.18	180	2	PVIN8
; -2442.03	767.03	180	2	PVIN8
; -2024.53	767.03	180	3	PVIN8
; -779.23	671.93	180	8	22
; -2533.23	598.93	180	1	23
; -2424.93	567.13	180	2	PVIN6
; -1182.73	441.83	180	6	SGND
; -2045.63	386.03	180	3	PVIN6
; -747.78	381.03	180	9	24
; -2553.23	321.13	180	1	PVIN6
; -2395.03	289.88	180	2	PVIN6
; -1696.23	253.88	180	5	PVIN6
; -2220.33	233.48	180	3	PVIN6
; -1958.33	218.38	180	4	25
; -912.13	118.08	180	7	SGND
; -2483.23	6.77	180	1	26
; -1977.43	-19.23	180	4	27
; -772.58	-53.33	180	9	28
; -772.53	-130.53	180	7	SGND
; -772.53	-211.68	180	8	29
; -2132.33	-221.93	180	3	30
; -912.13	-321.53	180	7	SGND
; -1970.28	-414.33	180	4	31
; -772.13	-598.43	180	8	32
; -1351.43	-742.68	180	6	33
; -2486.33	-790.23	180	1	34
; -1965.03	-929.63	180	4	35
; -2044.03	-1039.98	180	3	PGND
; -1695.43	-1128.28	180	5	PGND
; -2562.78	-1518.18	270	6	PGND
; -2415.43	-1362.03	270	7	PGND
; -2125.63	-1573.43	270	6	PGND
; -2030.88	-2279.58	270	2	PGND
; -2024.43	-1734.23	270	5	PGND
; -1759.28	-2059.93	270	4	PGND
; -1734.38	-1554.38	270	6	PGND
; -1594.33	-2566.78	270	1	PGND
; -1518.73	-1164.68	270	8	SGND
; -1382.23	-1344.48	270	7	38
; -1205.38	-2478.78	270	1	39
; -1173.63	-903.18	270	9	SGND
; -1087.93	-2036.08	270	3	PGND
; -800.73	-2012.08	270	4	40
; -800.73	-2459.38	270	1	41
; -772.13	-830.88	270	9	42
; -717.28	-1807.93	270	5	43
; -717.28	-2233.53	270	2	44
; -578.08	-2474.98	270	1	45
; -562.53	-2038.43	270	3	PVIN4
; -361.43	-2038.23	270	4	PVIN4
; -218.08	-2065.88	270	3	PVIN4
; -215.23	-2548.83	270	1	46
; -40.28	-2010.63	270	4	PVIN5
; -40.28	-2421.03	270	2	PVIN5
; -43.13	-2514.63	270	1	47
; 112.13	-2038.13	270	3	PVIN5
; 472.68	-2020.68	270	4	48
; 472.68	-2461.28	270	1	49
; 794.18	-1344.48	270	7	50
; 801.53	-2479.73	270	1	51
; 960.58	-1164.68	270	8	SGND
; 1110.48	-2037.93	270	3	PGND
; 1190.88	-1507.78	270	6	SGND
; 1341.98	-2013.98	270	4	PGND
; 1341.98	-2461.03	270	1	PGND
; 1367.78	-1396.78	270	7	52
; 1693.38	-2017.33	270	3	PGND
; 1685.08	-1607.88	270	6	PGND
; 1960.28	-2015.53	270	4	53
; 1960.33	-1197.73	270	8	54
; 2025.63	-1391.58	270	7	55
; 2025.98	-2310.53	270	2	PGND
; 2143.08	-2432.43	270	1	PGND
; 2399.38	-1653.73	270	6	PGND
; 2477.03	-2310.53	270	2	PGND
; 2477.28	-1391.58	270	7	56
; 2488.08	-1930.33	270	4	57
; 2396.98	-1036.23	0	2	PVIN7
; 2047.58	-989.03	0	3	PVIN7
; 6.38	-981.78	0	8	SGND
; 825.18	-864.98	0	7	70
; 2427.23	-813.13	0	2	PVIN7
; 2025.58	-813.13	0	3	PVIN7
; 2534.28	-799.28	0	1	71
; 2252.83	-621.48	0	2	PGND
; 1810.88	-620.08	0	4	PGND
; 2026.03	-591.68	0	3	PGND
; 825.18	-323.28	0	7	72
; 2046.58	-318.88	0	3	PGND
; 1190.88	-190.18	0	6	SGND
; 1367.78	-59.33	0	5	73
; 2488.98	-46.18	0	1	74
; 809.48	-27.78	0	7	75
; 956.08	156.83	0	6	SGND
; 2054.58	367.58	0	3	PGND
; 2488.38	392.18	0	1	76
; 809.48	400.23	0	7	77
; 809.48	627.13	0	7	78
; 2249.58	702.33	0	2	79
; 1815.83	702.88	0	4	80
; 2474.53	748.58	0	1	81
; 2024.43	748.58	0	3	82
; 798.53	979.33	0	7	SGND
; 9.63	979.33	0	8	SGND
; 2047.38	1017.23	0	3	PVIN2
; 2397.38	1130.33	0	2	PVIN2
; 1696.38	1153.83	0	4	PVIN2
; 1511.88	1187.33	0	5	SGND
; 2047.58	1317.38	0	3	PVIN2
; 2562.78	1536.33	90	7	96
; 2329.08	1589.13	90	7	PVIN2
; 2205.43	2489.03	90	3	PGND
; 1935.08	1549.58	90	7	PVIN2
; 1903.88	2054.13	90	6	PGND
; 1649.38	2054.43	90	5	PGND
; 1570.33	1965.88	90	6	97
; 1368.63	1375.23	90	8	98
; 1309.63	2034.03	90	6	99
; 1309.63	2483.23	90	3	100
; 945.53	2055.08	90	5	PVIN3
; 812.23	2430.18	90	3	PVIN3
; 812.23	2036.68	90	6	PVIN3
; 786.68	2541.93	90	2	1
; 672.88	2566.78	90	1	2
; 656.48	2114.33	90	5	PVIN1
; 503.13	2055.08	90	6	PVIN1
; 322.98	2054.13	90	5	PVIN1
; 81.17	2034.43	90	6	3
; 81.17	2488.83	90	3	4
; -314.13	2495.68	90	3	5
; -758.08	2054.13	90	6	PGND
; -786.98	1375.23	90	8	6
; -1002.03	2034.43	90	5	PGND
; -1002.03	2480.43	90	3	PGND
; -1178.73	1444.43	90	8	7
; -1196.43	1586.03	90	7	SGND
; -1296.28	2073.28	90	6	PGND
; -1529.68	2266.48	90	4	PGND
; -1556.98	1686.68	90	7	8
; -1941.78	1975.78	90	6	9
; -1980.98	2448.48	90	3	PGND
; -2049.38	2255.78	90	4	10
; -2099.68	1633.33	90	7	PVIN8O
; 3000	3000	45		
; )
; return data
; }
GetTestData()
{
data = 
(
-2396.98	1036.23	180	1	PVIN8O
-2047.58	989.03	180	4	PVIN8O
-6.38	981.78	180	7	SGND
-825.18	864.98	180	3	PVIN8
-2427.23	813.13	180	2	PVIN8
-2025.58	813.13	180	2	PVIN8
-2534.28	799.28	180	3	PVIN8
-2252.83	621.48	180	8	22
-1810.88	620.08	180	1	23
-2026.03	591.68	180	2	PVIN6
-825.18	323.28	180	6	SGND
-2046.58	318.88	180	3	PVIN6
-1190.88	190.18	180	9	24
-1367.78	59.33	180	1	PVIN6
-2488.98	46.18	180	2	PVIN6
-809.48	27.78	180	5	PVIN6
-956.08	-156.83	180	3	PVIN6
-2054.58	-367.58	180	4	25
-2488.38	-392.18	180	7	SGND
-809.48	-400.23	180	1	26
-809.48	-627.13	180	4	27
-2249.58	-702.33	180	9	28
-1815.83	-702.88	180	7	SGND
-2474.53	-748.58	180	8	29
-2024.43	-748.58	180	3	30
-798.53	-979.33	180	7	SGND
-9.63	-979.33	180	4	31
-2047.38	-1017.23	180	8	32
-2397.38	-1130.33	180	6	33
-1696.38	-1153.83	180	1	34
-1511.88	-1187.33	180	4	35
-2047.58	-1317.38	180	3	PGND
-2562.78	-1536.33	270	5	PGND
-2329.08	-1589.13	270	6	PGND
-2205.43	-2489.03	270	7	PGND
-1935.08	-1549.58	270	6	PGND
-1903.88	-2054.13	270	2	PGND
-1649.38	-2054.43	270	5	PGND
-1570.33	-1965.88	270	4	PGND
-1368.63	-1375.23	270	6	PGND
-1309.63	-2034.03	270	1	PGND
-1309.63	-2483.23	270	8	SGND
-945.53	-2055.08	270	7	38
-812.23	-2430.18	270	1	39
-812.23	-2036.68	270	9	SGND
-786.68	-2541.93	270	3	PGND
-672.88	-2566.78	270	4	40
-656.48	-2114.33	270	1	41
-503.13	-2055.08	270	9	42
-322.98	-2054.13	270	5	43
-81.17	-2034.43	270	2	44
-81.17	-2488.83	270	1	45
314.13	-2495.68	270	3	PVIN4
758.08	-2054.13	270	4	PVIN4
786.98	-1375.23	270	3	PVIN4
1002.03	-2034.43	270	1	46
1002.03	-2480.43	270	4	PVIN5
1178.73	-1444.43	270	2	PVIN5
1196.43	-1586.03	270	1	47
1296.28	-2073.28	270	3	PVIN5
1529.68	-2266.48	270	4	48
1556.98	-1686.68	270	1	49
1941.78	-1975.78	270	7	50
1980.98	-2448.48	270	1	51
2049.38	-2255.78	270	8	SGND
2099.68	-1633.33	270	3	PGND
2492.33	-1285.73	0	6	SGND
1971.33	-1190.98	0	4	PGND
936.23	-1187.33	0	1	PGND
2036.53	-1186.73	0	7	52
2427.33	-1112.18	0	3	PGND
2442.03	-767.03	0	6	PGND
2024.53	-767.03	0	4	53
779.23	-671.93	0	8	54
2533.23	-598.93	0	7	55
2424.93	-567.13	0	2	PGND
1182.73	-441.83	0	1	PGND
2045.63	-386.03	0	6	PGND
747.78	-381.03	0	2	PGND
2553.23	-321.13	0	7	56
2395.03	-289.88	0	4	57
1696.23	-253.88	0	2	PVIN7
2220.33	-233.48	0	3	PVIN7
1958.33	-218.38	0	8	SGND
912.13	-118.08	0	7	70
2483.23	-6.77	0	2	PVIN7
1977.43	19.23	0	3	PVIN7
772.58	53.33	0	1	71
772.53	130.53	0	2	PGND
772.53	211.68	0	4	PGND
2132.33	221.93	0	3	PGND
912.13	321.53	0	7	72
1970.28	414.33	0	3	PGND
772.13	598.43	0	6	SGND
1351.43	742.68	0	5	73
2486.33	790.23	0	1	74
1965.03	929.63	0	7	75
2044.03	1039.98	0	6	SGND
1695.43	1128.28	0	3	PGND
2562.78	1518.18	90	1	76
2415.43	1362.03	90	7	77
2125.63	1573.43	90	7	78
2030.88	2279.58	90	2	79
2024.43	1734.23	90	4	80
1759.28	2059.93	90	1	81
1734.38	1554.38	90	3	82
1594.33	2566.78	90	7	SGND
1518.73	1164.68	90	8	SGND
1382.23	1344.48	90	3	PVIN2
1205.38	2478.78	90	2	PVIN2
1173.63	903.18	90	4	PVIN2
1087.93	2036.08	90	5	SGND
800.73	2012.08	90	3	PVIN2
800.73	2459.38	90	7	96
772.13	830.88	90	7	PVIN2
717.28	1807.93	90	3	PGND
717.28	2233.53	90	7	PVIN2
578.08	2474.98	90	6	PGND
562.53	2038.43	90	5	PGND
361.43	2038.23	90	6	97
218.08	2065.88	90	8	98
215.23	2548.83	90	6	99
40.28	2010.63	90	3	100
40.28	2421.03	90	5	PVIN3
43.13	2514.63	90	3	PVIN3
-112.13	2038.13	90	6	PVIN3
-472.68	2020.68	90	2	1
-472.68	2461.28	90	1	2
-794.18	1344.48	90	5	PVIN1
-801.53	2479.73	90	6	PVIN1
-960.58	1164.68	90	5	PVIN1
-1110.48	2037.93	90	6	3
-1190.88	1507.78	90	3	4
-1341.98	2013.98	90	3	5
-1341.98	2461.03	90	6	PGND
-1367.78	1396.78	90	8	6
-1693.38	2017.33	90	5	PGND
-1685.08	1607.88	90	3	PGND
-1960.28	2015.53	90	8	7
-1960.33	1197.73	90	7	SGND
-2025.63	1391.58	90	6	PGND
-2025.98	2310.53	90	4	PGND
-2143.08	2432.43	90	7	8
-2399.38	1653.73	90	6	9
-2477.03	2310.53	90	3	PGND
-2477.28	1391.58	90	4	10
-2488.08	1930.33	90	7	PVIN8O
)
return data
}

GetTestData180()
{
data = 
(
-2492.33	1285.73	180	1	PVIN8O
-1971.33	1190.98	180	4	PVIN8O
-936.23	1187.33	180	7	SGND
-2036.53	1186.73	180	3	PVIN8
-2427.33	1112.18	180	2	PVIN8
-2442.03	767.03	180	2	PVIN8
-2024.53	767.03	180	3	PVIN8
-779.23	671.93	180	8	22
-2533.23	598.93	180	1	23
-2424.93	567.13	180	2	PVIN6
-1182.73	441.83	180	6	SGND
-2045.63	386.03	180	3	PVIN6
-747.78	381.03	180	9	24
-2553.23	321.13	180	1	PVIN6
-2395.03	289.88	180	2	PVIN6
-1696.23	253.88	180	5	PVIN6
-2220.33	233.48	180	3	PVIN6
-1958.33	218.38	180	4	25
-912.13	118.08	180	7	SGND
-2483.23	6.77	180	1	26
-1977.43	-19.23	180	4	27
-772.58	-53.33	180	9	28
-772.53	-130.53	180	7	SGND
-772.53	-211.68	180	8	29
-2132.33	-221.93	180	3	30
-912.13	-321.53	180	7	SGND
-1970.28	-414.33	180	4	31
-772.13	-598.43	180	8	32
-1351.43	-742.68	180	6	33
-2486.33	-790.23	180	1	34
-1965.03	-929.63	180	4	35
-2044.03	-1039.98	180	3	PGND
-1695.43	-1128.28	180	5	PGND
)
return data
}

GetTestData270()
{
data = 
(
-1285.73	-2492.33	270	1	PVIN8O
-1190.98	-1971.33	270	4	PVIN8O
-1187.33	-936.23	270	7	SGND
-1186.73	-2036.53	270	3	PVIN8
-1112.18	-2427.33	270	2	PVIN8
-767.03	-2442.03	270	2	PVIN8
-767.03	-2024.53	270	3	PVIN8
-671.93	-779.23	270	8	22
-598.93	-2533.23	270	1	23
-567.13	-2424.93	270	2	PVIN6
-441.83	-1182.73	270	6	SGND
-386.03	-2045.63	270	3	PVIN6
-381.03	-747.78	270	9	24
-321.13	-2553.23	270	1	PVIN6
-289.88	-2395.03	270	2	PVIN6
-253.88	-1696.23	270	5	PVIN6
-233.48	-2220.33	270	3	PVIN6
-218.38	-1958.33	270	4	25
-118.08	-912.13	270	7	SGND
-6.77	-2483.23	270	1	26
19.23	-1977.43	270	4	27
53.33	-772.58	270	9	28
130.53	-772.53	270	7	SGND
211.68	-772.53	270	8	29
221.93	-2132.33	270	3	30
321.53	-912.13	270	7	SGND
414.33	-1970.28	270	4	31
598.43	-772.13	270	8	32
742.68	-1351.43	270	6	33
790.23	-2486.33	270	1	34
929.63	-1965.03	270	4	35
1039.98	-2044.03	270	3	PGND
1128.28	-1695.43	270	5	PGND
)
return data
}

GetTestData0()
{
data = 
(
2492.33	-1285.73	0	1	PVIN8O
1971.33	-1190.98	0	4	PVIN8O
936.23	-1187.33	0	7	SGND
2036.53	-1186.73	0	3	PVIN8
2427.33	-1112.18	0	2	PVIN8
2442.03	-767.03	0	2	PVIN8
2024.53	-767.03	0	3	PVIN8
779.23	-671.93	0	8	22
2533.23	-598.93	0	1	23
2424.93	-567.13	0	2	PVIN6
1182.73	-441.83	0	6	SGND
2045.63	-386.03	0	3	PVIN6
747.78	-381.03	0	9	24
2553.23	-321.13	0	1	PVIN6
2395.03	-289.88	0	2	PVIN6
1696.23	-253.88	0	5	PVIN6
2220.33	-233.48	0	3	PVIN6
1958.33	-218.38	0	4	25
912.13	-118.08	0	7	SGND
2483.23	-6.77	0	1	26
1977.43	19.23	0	4	27
772.58	53.33	0	9	28
772.53	130.53	0	7	SGND
772.53	211.68	0	8	29
2132.33	221.93	0	3	30
912.13	321.53	0	7	SGND
1970.28	414.33	0	4	31
772.13	598.43	0	8	32
1351.43	742.68	0	6	33
2486.33	790.23	0	1	34
1965.03	929.63	0	4	35
2044.03	1039.98	0	3	PGND
1695.43	1128.28	0	5	PGND
)
return data
}

GetTestData90()
{
data = 
(
1285.73	2492.33	90	1	PVIN8O
1190.98	1971.33	90	4	PVIN8O
1187.33	936.23	90	7	SGND
1186.73	2036.53	90	3	PVIN8
1112.18	2427.33	90	2	PVIN8
767.03	2442.03	90	2	PVIN8
767.03	2024.53	90	3	PVIN8
671.93	779.23	90	8	22
598.93	2533.23	90	1	23
567.13	2424.93	90	2	PVIN6
441.83	1182.73	90	6	SGND
386.03	2045.63	90	3	PVIN6
381.03	747.78	90	9	24
321.13	2553.23	90	1	PVIN6
289.88	2395.03	90	2	PVIN6
253.88	1696.23	90	5	PVIN6
233.48	2220.33	90	3	PVIN6
218.38	1958.33	90	4	25
118.08	912.13	90	7	SGND
6.77	2483.23	90	1	26
-19.23	1977.43	90	4	27
-53.33	772.58	90	9	28
-130.53	772.53	90	7	SGND
-211.68	772.53	90	8	29
-221.93	2132.33	90	3	30
-321.53	912.13	90	7	SGND
-414.33	1970.28	90	4	31
-598.43	772.13	90	8	32
-742.68	1351.43	90	6	33
-790.23	2486.33	90	1	34
-929.63	1965.03	90	4	35
-1039.98	2044.03	90	3	PGND
-1128.28	1695.43	90	5	PGND
)
return data
}
*/


/*
GetTestData180()
{
data = 
(
-2396.98	1036.23	180	1	PVIN8O
-2047.58	989.03	180	4	PVIN8O
-6.38	981.78	180	7	SGND
-825.18	864.98	180	3	PVIN8
-2427.23	813.13	180	2	PVIN8
-2025.58	813.13	180	2	PVIN8
-2534.28	799.28	180	3	PVIN8
-2252.83	621.48	180	8	22
-1810.88	620.08	180	1	23
-2026.03	591.68	180	2	PVIN6
-825.18	323.28	180	6	SGND
-2046.58	318.88	180	3	PVIN6
-1190.88	190.18	180	9	24
-1367.78	59.33	180	1	PVIN6
-2488.98	46.18	180	2	PVIN6
-809.48	27.78	180	5	PVIN6
-956.08	-156.83	180	3	PVIN6
-2054.58	-367.58	180	4	25
-2488.38	-392.18	180	7	SGND
-809.48	-400.23	180	1	26
-809.48	-627.13	180	4	27
-2249.58	-702.33	180	9	28
-1815.83	-702.88	180	7	SGND
-2474.53	-748.58	180	8	29
-2024.43	-748.58	180	3	30
-798.53	-979.33	180	7	SGND
-9.63	-979.33	180	4	31
-2047.38	-1017.23	180	8	32
-2397.38	-1130.33	180	6	33
-1696.38	-1153.83	180	1	34
-1511.88	-1187.33	180	4	35
-2047.58	-1317.38	180	3	PGND
)
return data
}

GetTestData270()
{
data = 
(
-1036.23	-2396.98	270	1	gnd
-989.03	-2047.58	270	4	gnd
-981.78	-6.38	270	7	gnd
-864.98	-825.18	270	3	gnd
-813.13	-2427.23	270	2	gnd
-813.13	-2025.58	270	2	gnd
-799.28	-2534.28	270	3	gnd
-621.48	-2252.83	270	8	gnd
-620.08	-1810.88	270	1	gnd
-591.68	-2026.03	270	2	gnd
-323.28	-825.18	270	6	gnd
-318.88	-2046.58	270	3	gnd
-190.18	-1190.88	270	9	gnd
-59.33	-1367.78	270	1	gnd
-46.18	-2488.98	270	2	gnd
-27.78	-809.48	270	5	gnd
156.83	-956.08	270	3	gnd
367.58	-2054.58	270	4	gnd
392.18	-2488.38	270	7	gnd
400.23	-809.48	270	1	gnd
627.13	-809.48	270	4	gnd
702.33	-2249.58	270	9	gnd
702.88	-1815.83	270	7	gnd
748.58	-2474.53	270	8	gnd
748.58	-2024.43	270	3	gnd
979.33	-798.53	270	7	gnd
979.33	-9.63	270	4	gnd
1017.23	-2047.38	270	8	gnd
1130.33	-2397.38	270	6	gnd
1153.83	-1696.38	270	1	gnd
1187.33	-1511.88	270	4	gnd
1317.38	-2047.58	270	3	gnd
)
return data
}

GetTestData0()
{
data = 
(
2396.98	-1036.23	0	1	gnd
2047.58	-989.03	0	4	gnd
6.38	-981.78	0	7	gnd
825.18	-864.98	0	3	gnd
2427.23	-813.13	0	2	gnd
2025.58	-813.13	0	2	gnd
2534.28	-799.28	0	3	gnd
2252.83	-621.48	0	8	gnd
1810.88	-620.08	0	1	gnd
2026.03	-591.68	0	2	gnd
825.18	-323.28	0	6	gnd
2046.58	-318.88	0	3	gnd
1190.88	-190.18	0	9	gnd
1367.78	-59.33	0	1	gnd
2488.98	-46.18	0	2	gnd
809.48	-27.78	0	5	gnd
956.08	156.83	0	3	gnd
2054.58	367.58	0	4	gnd
2488.38	392.18	0	7	gnd
809.48	400.23	0	1	gnd
809.48	627.13	0	4	gnd
2249.58	702.33	0	9	gnd
1815.83	702.88	0	7	gnd
2474.53	748.58	0	8	gnd
2024.43	748.58	0	3	gnd
798.53	979.33	0	7	gnd
9.63	979.33	0	4	gnd
2047.38	1017.23	0	8	gnd
2397.38	1130.33	0	6	gnd
1696.38	1153.83	0	1	gnd
1511.88	1187.33	0	4	gnd
2047.58	1317.38	0	3	gnd
)
return data
}

GetTestData90()
{
data = 
(
1036.23	2396.98	90	1	gnd
989.03	2047.58	90	4	gnd
981.78	6.38	90	7	gnd
864.98	825.18	90	3	gnd
813.13	2427.23	90	2	gnd
813.13	2025.58	90	2	gnd
799.28	2534.28	90	3	gnd
621.48	2252.83	90	8	gnd
620.08	1810.88	90	1	gnd
591.68	2026.03	90	2	gnd
323.28	825.18	90	6	gnd
318.88	2046.58	90	3	gnd
190.18	1190.88	90	9	gnd
59.33	1367.78	90	1	gnd
46.18	2488.98	90	2	gnd
27.78	809.48	90	5	gnd
-156.83	956.08	90	3	gnd
-367.58	2054.58	90	4	gnd
-392.18	2488.38	90	7	gnd
-400.23	809.48	90	1	gnd
-627.13	809.48	90	4	gnd
-702.33	2249.58	90	9	gnd
-702.88	1815.83	90	7	gnd
-748.58	2474.53	90	8	gnd
-748.58	2024.43	90	3	gnd
-979.33	798.53	90	7	gnd
-979.33	9.63	90	4	gnd
-1017.23	2047.38	90	8	gnd
-1130.33	2397.38	90	6	gnd
-1153.83	1696.38	90	1	gnd
-1187.33	1511.88	90	4	gnd
-1317.38	2047.58	90	3	gnd
)
return data
}
