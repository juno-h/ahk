



Class autocad
{
	__new(mode)
	{
		try {
			this.acad 	:=  (mode == "active") ? ComObjActive("AutoCAD.Application")
						: 	(mode == "create") ? ComObjCreate("AutoCAD.Application") 
						: 	""
		} catch e {
			msgbox , 16 , error, please open AutoCAD ..
			Return 
		}
				
	}

	Visible(mode := true)
	{
		return this.acad.Visible 	:= true	
	}

	ThisDrawing()
	{
		this.Drawing := this.acad.ActiveDocument
		return this.Drawing
	}

	SelectionSetsClear()
	{
		for object in this.Drawing.SelectionSets
   			this.Drawing.SelectionSets.Item(object.name).delete
		return
	}

	SelectOnScreen()
	{
		this.SelectionSetsClear()
		this.Screen := this.Drawing.SelectionSets.Add("temp")
		this.Screen.SelectOnScreen  
		return this.Screen
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