Class autocad
{
	__new(mode)
	{
		try {
			this.acad 	:= 	(mode == "active") ? ComObjActive("AutoCAD.Application")
					: 	(mode == "create") ? ComObjCreate("AutoCAD.Application") 
					: 	""

		} catch e {
			msgbox , 16 , error, please open AutoCAD ..
			exit 
		}
	}

	Visible(mode := true)
	{
		this.acad.Visible 	:= mode			
		return this
	}
	
	ThisDrawing()
	{
		this.Drawing := this.acad.ActiveDocument
		return this.Drawing
	}
	SelectionSetsClear()
	{	
		try {
			for object in this.Drawing.SelectionSets
				this.Drawing.SelectionSets.Item(object.name).delete
		} catch e {
			WinActivate, AutoCAD
			send, {esc}						
			for object in this.Drawing.SelectionSets
				this.Drawing.SelectionSets.Item(object.name).delete
		}
		return this
	}

	SelectOnScreen(ItemName := "temp" , FilterType := "" , FilterData := "")
	{
		count := 0
		(ItemName = "temp") ? this.SelectionSetsClear() : ""
		SelectOnScreen:
		try
			this.Screen := this.Drawing.SelectionSets.Add(ItemName)
		catch e {			
			count++
			this.SelectionSetsClear()
			if (count > 10) 
			{
				msgbox , 16, error, 오류가 있습니다`n 관리자에게 문의 하세요 (SelectOnScreen)
				exit
			}
			goto, SelectOnScreen
		}	
		
		FilterType ? this.Screen.SelectOnScreen(FilterType , FilterData) : this.Screen.SelectOnScreen()
		
		; return this.Screen
		return this.Screen
	}
	
	setTextStyle(textStyle := "Standard")
	{	
		objTextStyle := this.Drawing.TextStyles.Add(textStyle)
		if (textStyle == "will 표준")
			objTextStyle.SetFont("@굴림", True, False, 0, 0)	; 굴림		
		this.Drawing.ActiveTextStyle := objTextStyle
		return this
	}

	addLayers( layer )
	{
		object := this.Drawing.Layers.Add(layer)  
		return object
	}
	LoadLinetypes( Linetype )
	{
		object := this.Drawing.Linetypes.Load(Linetype, "acad.lin")
		return object
	}
	; object.AddArc(Center, Radius, StartAngle, EndAngle)
	
	addArc( X , Y,  Radius , StartAngle , EndAngle , layer := "" , color := "")
	{
		object := this.Drawing.ModelSpace.AddArc(acdPoint([X , Y ,0.0]) , Radius , StartAngle , EndAngle)
  		layer ? this.CheckLayer(object , layer) : ""
  		object.color := color
  		object.Update
  		return  object
	}
	addCircle( X , Y,  Radius , layer := "" , color := "")
	{
		object := this.Drawing.ModelSpace.AddCircle(acdPoint([X , Y ,0.0]) , Radius)
  		layer ? this.CheckLayer(object , layer) : ""
  		object.color := color
  		object.Update
  		return  object
	}

	addPoint( X , Y, layer := "" )
	{
		object := this.Drawing.ModelSpace.AddPoint(acdPoint([X , Y ,0.0]))
  		layer ? this.CheckLayer(object , layer) : ""
  		object.Update
  		return  object
	}
	addPad( X , Y, w, h, layer := "" )
	{
		pt := ComObjArray(5, 8) ; VT_R8 = 5																																;  Com object 배열 생성 
		pt[0] := pt[2] :=  	X - w																																			;  X, Y 좌표 생성  
		pt[1] := pt[7] :=	Y - h																																						 
		pt[3] := pt[5] := 	Y + h																																							
		pt[4] := pt[6] := 	X + w		
		object 			:= this.Drawing.ModelSpace.AddLightWeightPolyline(pt)
		object.Closed 	:= True				; 객체 닫기		
  		layer ? this.CheckLayer(object , layer) : ""
  		object.Update
  		return  object
	}
	addLwline(array , layer := "" , color := "" , ConstantWidth := "0" , Closed := "false" , re := "")
	{
		static before := ""
		arrayCount := array.count()
		pt := ComObjArray(5, array.count()) 
		loop % arrayCount
			pt[A_index - 1] := array[A_index]
		object := this.Drawing.ModelSpace.AddLightWeightPolyline(pt)		
  		layer ? this.CheckLayer(object , layer) : ""
		object.ConstantWidth 	:= ConstantWidth		; 선 두께
		object.Closed 		:= Closed			; 객체 닫기
		object.color 		:= color
		before 			:= object
  		object.Update

  		return  re ? {object : object , result : result} : object
	}
	TranslateCoordinates(object, array)
	{
		pt := ComObjArray(5, array.count()) 
		loop % array.count()
			pt[A_index - 1] := array[A_index]
		object.Coordinates := pt
		return this
	}
	addText(X , Y, text , size := 10 , StyleName := "" , layer := "" , color := "")
	{
		pos := acdPoint([X , Y ,0.0])
		object := this.Drawing.ModelSpace.AddText(text, pos, size)
  		layer ? this.CheckLayer(object , layer) : ""
		object.Alignment		:=	"10"
		object.TextAlignmentPoint	:=	pos
		object.color			:=	color
		try
			object.StyleName		:=	StyleName
  		object.Update
  		return  object
	}
	; object base  example , obj := {x : obj.x[1] , y : obj.y[1] , text : type , size : 50 , layer : step . "F " , angle : temp.angle , textgap : 35 }
	addTextAlign(obj)				; channel text 생성 시 적용
	{
		
		varAlignA := obj.mode ? 11 : 9
		varAlignB := obj.mode ? 9  : 11
		
			(obj.angle =   0) ? ( obj.x := obj.x + obj.textgap , Alignment := varAlignA , Rotation := 0		)
		 : 	(obj.angle =  90) ? ( obj.y := obj.y + obj.textgap , Alignment := varAlignA , Rotation := 1.570796	)
		 : 	(obj.angle = 180) ? ( obj.x := obj.x - obj.textgap , Alignment := varAlignB , Rotation := 0		)
		 : 	(obj.angle = 270) ? ( obj.y := obj.y - obj.textgap , Alignment := varAlignB , Rotation := 1.570796	) 
		 : ""
		
		pos 			:=	acdPoint([obj.x , obj.y ,0.0])
		object 			:=	this.Drawing.ModelSpace.AddText(obj.text, pos, obj.size)
		object.layer	:=	obj.layer ? this.CheckLayer(object , layer) 	: "0"
		object.Height					:=	obj.size
		object.Rotation					:=	Rotation
		object.Alignment				:=	Alignment
		object.TextAlignmentPoint		:=	pos
  		object.Update
  		return  object
	}
	addHatch(entity)
	{
		OuterLoop 	:= ComObjArray(9, 1)
		OuterLoop[0] 	:= entity
		AChatch := this.acad.ActiveDocument.ModelSpace.AddHatch(1, "SOLID", true)
		try
		{
			AChatch.AppendOuterLoop(OuterLoop)
			AChatch.Evaluate
			AChatch.Update		
		}
		return this
	}
	delete(object)
	{
		object.Delete
  		return  object
	}
	move(object , point1 , point2)
	{
		try
			object.move( acdPoint(point1) , acdPoint(point2) )
		catch e 
			object.move( point1 , acdPoint(point2) )
		return object
	}
	CheckLayer(object , layer)
	{
		try 
			object.layer := layer
		catch e {
			this.addLayers( layer )
			object.layer := layer
		}
		return this
	}
	setprompt(content)
	{
		this.Drawing.Utility.prompt(content)
		return this
	}
	getPoint()
	{
		temp := []
		CadScreen := this.SelectOnScreen(setComArray(VT_I2 := 2, ["0"]) , setComArray(VT_VARIANT := 0xC, ["Point"]))		
		for object in CadScreen
		{
			point := object.Coordinates
			temp.push(point[0] , point[1])
		}
		return temp
	}
	getClickPoint()
	{
		temp := []
		loop {
			try {
				Point := this.Drawing.Utility.GetPoint(,A_index " 번째 위치를 지정하세요")
				temp.push(Point[0] , Point[1])
			} catch e {
				Point := this.Drawing.Utility.prompt(A_index " 번째는 취소 되었습니다.`n")
				break
			}
		}
		return temp
	}
	getScreenClickPoint()
	{
		try
			return this.acad.ActiveDocument.Utility.GetPoint
		catch e
		{
			msgbox , 16 , ERROR , EXIT
			exit
		}
	}
	SetString(String)
	{
		object.TextString := String
		object.update
		return this
	}

}

linePoint(Screen, objstatus := "")
{
	array := []
	for object in Screen	{ 
		a := A_index
		if InStr(object.ObjectName, "Polyline")
		{
			point := []
			for pos in object.Coordinates {
				mod(A_index,2) 
				? point["x" , point["x"].count() ? point["x"].count()+1 : 1] := pos
				: point["y" , point["y"].count() ? point["y"].count()+1 : 1] := pos	

			}			
		}
		array[a] := {	x 		: point["x"] 
					, 	y 		: point["y"] 
					, 	maxX 	: max(point["x"]*)
					, 	maxY 	: max(point["y"]*)
					, 	minX 	: min(point["x"]*)
					, 	minY 	: min(point["y"]*)
					, 	wid  	: (max(point["x"]*)-min(point["x"]*))
					, 	hei  	: (max(point["y"]*)-min(point["y"]*))	}

		objstatus ? objstatus.setContent("cad " a, 1)
		.setContent(array[a].maxX " . "
		.	array[a].maxY " . "
		.	array[a].wid " . "
		.	array[a].hei " ... "
		,  2) : ""
	}
	return array
}
acdPoint(Values)
{
	arr := ComObjArray(VT_R8 := 5, Values.count())
	for i, v in Values
		arr[i-1] := v
	return arr
}
setComArray(Type, Values)
{
	; Type=0xC
	arr := ComObjArray(Type, Values.count())
	for i, v in Values
		arr[i-1] := v
	return arr
}
