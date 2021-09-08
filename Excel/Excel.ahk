Class excel
{
	__new(mode)
	{
		this.cell := ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]

		if (mode == "active"){
			try
				this.xl 	:=  ComObjActive("Excel.Application")
			catch e {
				msgbox , 16 , error, please open excel ..
				exit 
			}
		} else if (mode == "create"){
			try
				this.xl 	:=  ComObjCreate("Excel.Application")
			catch e {
				msgbox , 16 , error, please open excel ..
				exit 
			}
			
			ComObjError(false)
		}		
	}
	Visible(var)
	{
		this.xl.Visible := var
		return this
	}
	Workbooks_open(path)
	{
		this.workbook := this.xl.WorkBooks.Open(path)
		return this
	}
	Workbooks_save()
	{
		this.workbook.save
		return this
	}
	Workbooks_add()
	{
		this.xl.WorkBooks.add
		return this
	}
	Workbooks_saveas(filname)
	{
		this.xl.ActiveWorkbook.SaveAs(filname , 56)             ;51 is an xlsx, 56 is an xls
		Sleep,500
		this.xl.ActiveWorkbook.Save
		return this
	}
	Workbooks_close()
	{
		this.xl.WorkBooks.Close()
		this.xl.quit
		return this
	}
	Select(area)
	{	
		this.xl.Range( area ).Select
		return
	}
	SelectionCopy( mode := "")
	{	
		if (mode == "all") {
			cell := this.checkdata()
			this.xl.Range(cell).CurrentRegion.Select
		}
		try {
			this.xl.Selection.Copy
			if !Strlen(regexreplace(Clipboard , "[!`n`r]"))
				return false
			this.data 		:= Clipboard
			Clipboard 		:= ""
		} catch e{
			msgbox , 16 , error, please select area ..
			exit
		}
		return this.data
	}
	checkdata(mode := "")
	{
		try {
			AtvCell 	:= this.xl.ActiveCell.Address	
			AtvCellCheck 	:= this.xl.Range(AtvCell).value
		} catch e {
			msgbox , 16 , error, please open excel ..
			exit 
		}
		if (AtvCellCheck = "") {	
			msgbox , 16 , error, 선택된 셀의 데이터가 유효하지 않습니다.`n다시 선택해주세요.
			exit
		}
		return RegExReplace(AtvCell,"[^A-Z,^0-9]")
	}
	getCellbyNumber(number)
	{		
		static hr 
		hr 		:= floor((number + hr) / 27)
		tail 	:= mod(number+hr , 27)
		return this.cell[hr] . this.cell[tail]
	}
	getSelectionAddress()
	{
		try
			Address := this.xl.Selection.Address
		return Address
	}
	getActiveCell()
	{
		try
			Address := this.xl.ActiveCell.Address
		return Address
	}
	getActiveCellRowNumber()
	{
		cell := this.xl.ActiveCell.Address
		return RegExReplace(cell, "[^0-9]")
	}
	getActiveCellColName()
	{
		cell := this.xl.ActiveCell.Address
		return RegExReplace(StrReplace(cell,"$"), "[^A-Z]")
	}
	getlineDataOnlyStartNumber()
	{
		array := []	
		loop , parse, % this.data ,`n ,`r
		{
			field := StrSplit(A_loopfield,"`t")
			check := field[1]
			if check is not number 
				continue
			array.push(A_loopfield)
		}
		return array
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
					if (A_index = field.count() and ((field.count()+array_stay_count) <> Columnscount))
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
	getColumnWidth(column)
	{
		return this.xl.Columns(column).ColumnWidth
	}
	getUsedRage()
	{
		return this.xl.ActiveSheet.usedRange.Address()
	}
	getColumnAdd()
	{
		cell_first 	:= 0
		cellcount 	:= 0
		firstadd 	:= false		
		add 		:= StrSplit(RegExReplace(this.xl.ActiveSheet.UsedRange.Address,"[^A-Z,:]"),":")		
		loop , 26
		{	
			cell_second := 0
			loop , 26
			{
				cellcount++
				cell_second++
				cellname := this.cell[cell_first] . this.cell[cell_second]
				if (add[1] = cellname and firstadd = false)
				{
					StartNumber := (cell_first*26) + cell_second
					firstadd 	:= true
					cell_first 	:= 0
					cell_second := 0					
				} else if (add[2] = cellname)
				{
					break
				}
			}
			if (add[2] = cellname)
			{
				break
			}
			cell_first++
		}
		if ((cellcount-1) > 26)
		{
			msgbox , 16 ,error,Check !!! method - getColumnAdd of Class excel  
			exit
		}
		temp 				:= []
		temp["Usedcell"] 	:= []
		index 				:= StartNumber-1
		loop % cellcount-1
		{
			temp["Usedcell"].push(this.cell[index+=1])
		}
		temp["startcol"] 	:= add[1]
		temp["endcol"] 		:= add[2]
		temp["count"] 		:= cellcount-1
		return temp
	}
	; control	──────────────────────────────────────────────────────────────────────────────────────
	setoffset(RowOffset , ColumnOffset)
	{
		return this.xl.ActiveCell.Offset(RowOffset, ColumnOffset).Select
	}
	setActiveateCell(area)
	{
		return this.xl.Range( area ).Activate
	}
	setLineStyleAll(cell , lineStyle , weight , color := "")
	{
		entity := this.xl.range( cell )		
		entity.Borders.LineStyle 	:= lineStyle
		entity.Borders.weight 		:= weight
		color ? (entity.Border.color := color) : ""
		return this
	}
	; font 	──────────────────────────────────────────────────────────────────────────────────────
	setFontName(area , font)
	{
		this.xl.Range( area ).Font.name := font
		return this
	}
	setFontSize(area , size)
	{
		this.xl.Range( area ).Font.size := size
		return this
	}
	setFontColor(area , Color)
	{
		; null - 0 , black - 1 , white - 2 , red - 3 , green - 4 , blue - 5 , yellow - 6 , etc seaech vba color index
		this.xl.Range( area ).Font.ColorIndex := Color
		return this
	}
	; font 	──────────────────────────────────────────────────────────────────────────────────────
	setCellColor(area , ColorIndex)
	{
		; null - 0 , black - 1 , white - 2 , red - 3 , green - 4 , blue - 5 , yellow - 6 , etc seaech vba color index
		try
			this.xl.Range( area ).Interior.ColorIndex := ColorIndex
		return this
	}
	setColumnWidth(column , size)
	{
		this.xl.Columns(column).ColumnWidth := size
		return this
	}
	addSheet()
	{
		this.xl.Sheets.Add()
		; this.xl.Sheets.Add(_, this.xl.Sheets(this.xl.Sheets.Count)) 
		return thiss
	}
	paste(cell , value)
	{
		Clipboard := value
		Clipwait

		try {
			this.xl.Range(cell).pasteSpecial paste 
			Clipboard := ""
		} Catch e {
			; msgbox , 16, error, % e.message "`n cell : "  cell "`nvalue : " value
			return false
		}
		
		return this
	}

}
