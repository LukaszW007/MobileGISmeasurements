Option Explicit

dim Debug, DebugGPS, debugPow, objPlik
dim objAPTimer 
dim objPlikShape
dim NrProducenta, Wlasciciel, Inspektor, Data

dim Pi 
dim GPSInterval 
dim GPSMaska

dim GPS_DM, GPS_EM, GPSstop, objFile, Laczenie,objPunkty,NowaLinia 

dim GPSnmeaTryb, Czas, NowaPozycja, X,Y,Z,XC,YC,ZC,X0,Y0,Z0 

dim Pomiar, Ciagly, GPSstatic 
dim OffsetTyp,Offset1Az,Offset1Dist,Offset2Dist,OffsetPamietaj

dim Aktualizacja 
dim Powierzchnia 
dim PowierzchniaTyp 
dim PowierzchniaWewNowa 
dim PowierzchniaZewNowa 
dim PowierzchniaBlokada 
dim Pwew, Pzew,Owew,Ozew
dim PPwewX, PPwewY,OPwewX,OPwewY,PPzewX,PPzewY,OPzewX,OPzewY
dim PwewXmin 
dim PzewXmin 
dim PPX, PPY

dim wsWyniki 
dim wsPha, wsPhaWew 'Pow w ha
dim wsPm2, wsPm2Wew'Powi w m2
dim wsO, wsOWew  'Obw

dim wsX, wsY 'Współrzędne wybranego pkt
dim wsPLB 'Typ rysowanego obiektu
dim wsWskazany 'Czy odległość i azymut liczyć od wskazanego punktu czy od ostatnio pomierzonego

dim kPow ' sumarycza powierzchnia w obliczeniach kalkulatorowych
dim kObw ' sumaryczny obwód w obliczeniach kalkulatorowych

dim tempPow
dim tempObw


Dim g_pForm	
Dim g_pNowyPageControls  
Dim g_pKonfPageControls  
Dim g_pDanePageControls
Dim g_pDane2PageControls
Dim wsRej 



Sub InitForm2
'Called when form loads 

	'Get references to the form and each page's controls
	Set g_pForm = ThisEvent.Object
	'Set g_pDane2PageControls = g_pForm.Pages("Page2").Controls
	Set g_pKonfPageControls = g_pForm.Pages("Page1").Controls
End Sub

Sub CleanUp2
'Called when form unloads

	'Free resources
	Set g_pForm = Nothing
	Set g_pKonfPageControls = Nothing
	set g_pDane2PageControls = nothing
End Sub

Sub InitForm
'Called when form loads 

	'Get references to the form and each page's controls
	Set g_pForm = ThisEvent.Object
	Set g_pNowyPageControls = g_pForm.Pages("page1").Controls
	Set g_pDanePageControls = g_pForm.Pages("PAGE2").Controls
End Sub

Sub CleanUp
'Called when form unloads

	'Free resources
	Set g_pForm = Nothing
	Set g_pNowyPageControls = Nothing
	set g_pDanePageControls = nothing
End Sub


sub NowaMapa
dim objForm
dim no
	no=msgbox ("Czy utworzyć nowy obiekt",vbYesNo,"IACS")
	if no=6 then
		Set objForm = Applet.Forms.item("form1")
		objForm.show
	end if
end sub

Sub ShowDirectoryBrowser
'Called when directory browser button is clicked
	
	'Prompt the user for a file in the folder of interest
	Dim Folder
	Folder = CommonDialog.ShowOpen(,,"Wskaż dowolny plik w folderze przeznaczenia", &H800)
	If Not IsEmpty(Folder) Then
		g_pNowyPageControls("Edit2").Value = pobierzfolder(Folder)
	End If
End Sub

Function pobierzfolder(p_sciezkaPliku)
'Utility function to return the folder only given a complete file path

	'Get position of last backslash
	Dim intLastBSPos
	intLastBSPos = InStrRev(p_sciezkaPliku,"\")

	
	pobierzfolder = Left(p_sciezkaPliku,intLastBSPos-1)
End Function

sub cmdNazwaProducentaOK
dim a, p_sciezkaPliku, b, nl
dim objPlikShape, katalog, wynik
Dim UkWsp, UkWspFile
	if not NrProducenta="" then
			nl=msgbox ("Czy na pewno chcesz zamknąć bieżącą robotę?",vbYesNo,"Zamykanie")
		if nl=6 then
		Application.Map.Layers.Remove(NrProducenta&"punkty.shp") 
		Application.Map.Layers.Remove(NrProducenta&"linie.shp") 
		Application.Map.Layers.Remove(NrProducenta&"powierzchnie.shp")
		Application.Map.Refresh
		NrProducenta=""
		else
		exit sub
		end if
	end if
	set objPlikShape=CreateAppObject("FILE")
	set a=ThisEvent.Object.Pages(1).Controls
	set b=ThisEvent.Object.Pages(2).Controls
	
	NrProducenta=a("Edit1")
	Wlasciciel=b("Edit1")
	Inspektor=b("Edit2")
	Data=b("Date1")
	
	katalog= a("Edit2")
	katalog=katalog&"\"&NrProducenta
	wynik=objPlikShape.CreateDirectory ( katalog )
	katalog=katalog&"\"&cstr(NrProducenta)
	If not wynik then
		msgbox "Katalog o tej nazwie już istnieje."&vbcrlf&"Otwieram istniejaca mape.",vbOkOnly,"Otwieranie"
		
		Call Map.AddLayerFromFile (katalog&"powierzchnie.shp")
		Call Map.AddLayerFromFile (katalog&"linie.shp")
		Call Map.AddLayerFromFile (katalog&"punkty.shp")
		For Each objPlikShape in Map.Layers
			if objPlikShape.Name = (NrProducenta&"punkty.shp") then
				objPlikShape.Editable=true
			end if
			if objPlikShape.Name = (NrProducenta&"linie.shp") then
				objPlikShape.Editable=true
			end if
			if objPlikShape.Name = (NrProducenta&"powierzchnie.shp") then
				objPlikShape.Editable=true
			end if
		next
		exit sub
	end if
	
	Set UkWsp = Application.CreateAppObject("CoordSys")
	UkWspFile=Application.Path&"\Applets\UkWsp\WGS84.prj"
	UkWsp.Import UkWspFile
	Set Map.CoordinateSystem = UkWsp


	'poligon
	Set objPlikShape = Application.CreateAppObject ("RecordSet")
	Call objPlikShape.Create (katalog&"powierzchnie.shp" , apShapePolygon)
	Call objPlikShape.Fields.Append ("Oznaczenie", apFieldCharacter)
	Call objPlikShape.Fields.Append ("NrEwid", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Uprawa", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Grupa upr", apFieldCharacter)
	Call objPlikShape.Fields.Append ("NrProd", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Producent", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Inspektor", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Data pom", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Pow0", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Obw0", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Uwagi", apFieldCharacter)
	Call Map.AddLayerFromFile (katalog&"powierzchnie.shp")

	'polilinia
	Set objPlikShape = Application.CreateAppObject ("RecordSet")
	Call objPlikShape.Create (katalog&"linie.shp" , apShapePolyline)
	Call objPlikShape.Fields.Append ("Oznaczenie", apFieldCharacter)
	Call objPlikShape.Fields.Append ("NrEwid", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Uprawa", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Grupa upr", apFieldCharacter)
	Call objPlikShape.Fields.Append ("NrProd", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Producent", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Inspektor", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Pow0", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Obw0", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Data pom", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Uwagi", apFieldCharacter)
	Call Map.AddLayerFromFile (katalog&"linie.shp")

	'punkt
	Set objPlikShape = Application.CreateAppObject ("RecordSet")
	Call objPlikShape.Create (katalog&"punkty.shp" , apShapePoint)
	Call objPlikShape.Fields.Append ("Oznaczenie", apFieldCharacter)
	Call objPlikShape.Fields.Append ("NrEwid", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Uprawa", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Grupa upr", apFieldCharacter)
	Call objPlikShape.Fields.Append ("NrProd", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Producent", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Inspektor", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Pow0", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Obw0", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Data pom", apFieldCharacter)
	Call objPlikShape.Fields.Append ("Uwagi", apFieldCharacter)
	Call Map.AddLayerFromFile (katalog&"punkty.shp")


	For Each objPlikShape in Map.Layers
		if objPlikShape.Name = (NrProducenta&"punkty.shp") then
			objPlikShape.Editable=true
		end if
	next

	For Each objPlikShape in Map.Layers
		if objPlikShape.Name = (NrProducenta&"linie.shp") then
			objPlikShape.Editable=true
		end if
	next

	For Each objPlikShape in Map.Layers
		if objPlikShape.Name = (NrProducenta&"powierzchnie.shp") then
			objPlikShape.Editable=true
		end if
	next
	
end sub

Sub LoadSelectedShapefile
	Dim Folder, filtrPliku

  filtrPliku = "ArcPad Layers|*.shp"
	Folder = CommonDialog.ShowOpen(,filtrPliku,"Wskaż plik podkładu ewidencyjnego", &H800)
	If Not IsEmpty(Folder) Then
		g_pNowyPageControls("Edit3").Value = pobierzfolder(Folder)
		Application.Map.AddLayerFromFile (Folder)  
    	Application.Map.Refresh

	End If

End Sub

Function pobierzfolder(p_sciezkaPliku)
'Utility function to return the folder only given a complete file path

	'Get position of last backslash
	Dim intLastBSPos
	intLastBSPos = InStrRev(p_sciezkaPliku,"\")

	
	pobierzfolder = Left(p_sciezkaPliku,intLastBSPos-1)
End Function

Sub UnLoadSelectedShapefile
'Called when Load Shapefile button is clicked
dim nl
	nl=msgbox ("Czy na pewno chcesz zamknąć bieżącą robotę?",vbYesNo,"Zamykanie")
	if nl=6 then
		Application.Map.Layers.Remove(NrProducenta&"punkty.shp") 
		Application.Map.Layers.Remove(NrProducenta&"linie.shp") 
		Application.Map.Layers.Remove(NrProducenta&"powierzchnie.shp")
		Application.Map.Refresh
		NrProducenta=""
		Application.Map.clear
	else
		exit sub
	end if
		
		
End Sub

sub GPS_NowaPozycja
	NowaPozycja=true
end sub


sub OnTimer
dim i,l 
	
	objAPTimer.Enabled =false

	if not(GPSstatic) then

		if GPSnmeaTryb=9 then 
			application.gps.write ("em,,/msg/nmea/ZDA"+GPSIntervals+vbcr+vblf)
			GPSnmeaTryb=0
		end if

		if GPSnmeaTryb=8 then 
			application.gps.write ("em,,/msg/nmea/VTG"+GPSIntervals+vbcr+vblf)
			GPSnmeaTryb=9
		end if

		if GPSnmeaTryb=7 then 
			application.gps.write ("em,,/msg/nmea/RMC"+GPSIntervals+vbcr+vblf)
			GPSnmeaTryb=8
		end if

		if GPSnmeaTryb=6 then 
			application.gps.write ("em,,/msg/nmea/GSV"+GPSIntervals+vbcr+vblf)
			GPSnmeaTryb=7
		end if

		if GPSnmeaTryb=5 then 
			application.gps.write ("em,,/msg/nmea/GSA"+GPSIntervals+vbcr+vblf)
			GPSnmeaTryb=6
		end if

		if GPSnmeaTryb=4 then 
			application.gps.write ("em,,/msg/nmea/GLL"+GPSIntervals+vbcr+vblf)
			GPSnmeaTryb=5
		end if

		if GPSnmeaTryb=3 then 
			application.gps.write ("em,,/msg/nmea/GGA"+GPSIntervals+vbcr+vblf)
			GPSnmeaTryb=4
		end if

		if GPSnmeaTryb=2 then 
			application.gps.write ("set,/par/pos/glo/sat,off"+vbcr+vblf)
			GPSnmeaTryb=3
		end if


		if GPSnmeaTryb=1 then 
			application.gps.write ("dm"+vbcr+vblf)
			GPSnmeaTryb=2
		end if

	end if
	
	

	if  (pomiar) then 
		call PomiarPunktu
		call RysowaniePunktu
		if powierzchnia then call ObliczPowierzchnie
		call PokazWyniki
	end if
	
	objAPTimer.Interval =1000
	objAPTimer.Enabled =true

end sub

sub test(a)
dim sl
dim rs
dim sh, shtype
dim part
dim vertex
dim i, j

dim b

	tempPow=0
	tempObw=0

	set sl=Application.Map.SelectionLayer
	if sl is nothing then exit sub
	
	set rs=sl.Records
	rs.bookmark=Application.Map.SelectionBookmark
	
	set sh=rs.Fields.Shape
	shtype=sh.ShapeType
	
	if shtype>10 then shtype=shtype-10
	if shtype>10 then shtype=shtype-10

	'shtype 1=point, 3=polyline, 5=polygon
	
	if shtype=1 then exit sub
	if shtype=3 then exit sub
	
	j=1
	for each part in sh.parts
		i=1
		for each vertex in part
			'msgbox cstr(vertex.x)&" "&cstr(vertex.y),0,cstr(i)&" "&cstr(j)
			i=i+1
		next
		j=j+1
	next

	'Usuniecie zaznaczonego obiektu
	'rs.Delete 
	
	'Powierzchnia
	tempPow=abs(sh.area)


	'Obwód
	tempObw=abs(sh.perimeter)

	'Właściwości
	'rs.Fields("Pow0").value=cstr(int(tempPow))
	'rs.Fields("Obw0").value=cstr(int(tempObw))
	'rs.Update

end sub

sub OnFeatureChanged
end sub

sub cmdPoleObwod
Dim a, p
dim st

	call test(1)
	p=abs(tempObw)
	a=abs(tempPow)

	msgbox "Powierzchnia= "&formatnumber(a,0)&" [m2]"&vbcr&vblf& _
"Tolerancja  = "&formatnumber(1.5*p,0)&" [m2]"&vbcr&vblf& _
"Powierzchnia= "&formatnumber(0.0001*a,2)&" [ha]"&vbcr&vblf& _
"Tolerancja  = "&formatnumber(0.0001*1.5*p,2)&" [ha]"&vbcr&vblf& _
"Obwod= "&formatnumber(p,0)&" [m]",vbOkOnly,"Powierzchnia/Obwod"
end sub


sub ObliczenieOffsetu
dim dx,dy,d,temp

dx=0
dy=0
d=0

select case OffsetTyp
	case 1
		temp=pi/2-Offset1Az
		dx=Offset1Dist*cos(temp)
		dy=Offset1Dist*sin(temp)
		if DebugGPS then msgbox cstr(dx)&" "&cstr(dy),0,"dx, dy"
	case 2
		dx=X-X0
		dy=Y-Y0
		d=sqr(dx*dx+dy*dy)
		if d<0.01 then d=0.01 'Tak na wszelki wypadek żeby nie dzielić przez zero
		dx=-Offset2Dist*dx/d
		dy=Offset2Dist*dy/d
		if DebugGPS then msgbox cstr(d)&" "&cstr(dx)&" "&cstr(dy),0,"d dx, dy"
		temp=dx
		dx=dy
		dy=temp
end select

XC=X+dx
YC=Y+dy
ZC=Z

If not(OffsetPamietaj) then OffsetTyp=0

end sub
sub PokazWyniki
dim t, s, dx, dy
dim p
	
	if not(wsWyniki) then exit sub

	p = Map.CoordinateSystem.Projection '43000
	
	
	For each T in Application.Toolbars

		if t.name="tlbInfo" then
			
			t.Visible = true

			s=""
			if wsPha then s=s&"P="&FormatNumber(0.0001*(abs(Pzew)-abs(Pwew)),4)&" [ha]; "
			if wsPm2 then s=s&"P="&FormatNumber((abs(Pzew)-abs(Pwew)),0)&" [m2]; "
			
			if wsO then s=s&"O="&FormatNumber(Ozew,0)&" [m]; "
			
			if wsPhaWew then s=s&"Pwew="&FormatNumber(0.0001*abs(Pwew),4)&" [ha]; "
			if wsPm2Wew then s=s&"Pwew="&FormatNumber(abs(Pwew),0)&" [m2]; "
			
			if wsOWew then s=s&"Owew="&FormatNumber(Owew,0)&" [m]; "
			
			if not(wsWskazany) then
				dx=wsX-Xc
				dy=wsY-Yc
			else
				dx=Xc-X0
				dy=Yc-Y0
			end if
			
					
			t.caption=s
			exit sub

		end if

	Next


end sub
Sub PomiarPunktu
dim temp, w, i, j

	if (DebugGPS) or ((GPS.IsOpen) and (NowaPozycja)) then
		X0=X
		Y0=Y
		Z0=Z
		Czas=now()
		if debugGPS then
			temp=objPlik.Readline
			i=instr(1,temp,vbTab)
			j=instr(i+1,temp,vbTab)
			w=mid(temp,1,i-1)
			X=cdbl(w)
			if j=0 then j=len(temp)
			w=mid(temp,i+1,j-i)
			Y=cdbl(w)
			if len(temp)=j then
				Z=cdbl(0.0)
			else
				w=mid(temp,j+1,len(temp)-j)
				z=cdbl(w)
			end if
			nowaPozycja=true
			StatusBar.Text(1) = cstr(x)&"*"&cstr(y)&"*"&cstr(z)
		else
			X=GPS.X
			Y=GPS.Y
			Z=GPS.Z
			nowaPozycja=false
		end if
		
		call ObliczenieOffsetu
		if not(Ciagly) then Pomiar=False
	else
		if not(GPS.isopen) then 
			call PortGPSZamkniety
		end if
	end if

end sub

sub Timer1
	objAPTimer.Interval =1000
	objAPTimer.Enabled =true
end sub

sub Timer0
	objAPTimer.Enabled =false
end sub

function GPSIntervals
	if isnumeric(GPSInterval) then
		GPSIntervals=":"&cstr(GPSInterval)
	else
		GPSInterval=1
		GPSIntervals=":"&cstr(GPSInterval)
	end if
end function

sub PortGPSZamkniety
	msgbox "Port GPS Zamknięty",vbOKonly,"Uwaga!"
	pomiar=false
end sub

sub GPS_NowaPozycja
	NowaPozycja=true
end sub

sub RysowaniePunktu
Dim objPunkt, nl, objSymbol, objLLinia, objLPunkty, objLPunkt

set objPunkt=Application.CreateAppObject("Point")
objPunkt.X=XC
objPunkt.Y=YC
objPunkt.Z=ZC

if (wsPLB="Punkt") or (PowierzchniaBlokada) then
		If Not Application.Map.AddFeature (objPunkt,false) Then
			MsgBox "Punktu nie dodano",vbExclamation,"Błąd"
			Pomiar=false
		End If
end if

if (wsPLB="Linia") and (not(PowierzchniaBlokada)) then

		nl=0
		if NowaLinia then
			nl=msgbox ("Czy rozpocząć rysowanie nowej linii",vbYesNo,"Linie")
			if nl=6 then
				set objPunkty=Application.CreateAppObject("Points")
			end if
		end if
		NowaLinia=false

		If Not Application.Map.AddFeature (objPunkt,false) Then
			MsgBox "Punktu nie dodano",vbExclamation,"Błąd"
			Pomiar=false
		End If		
	
		objPunkty.add objPunkt

end if

if (wsPLB="Powierzchnia") and (not(PowierzchniaBlokada)) then

		nl=0
		if NowaLinia then
			nl=msgbox ("Czy rozpocząć rysowanie nowej powierzchni",vbYesNo,"Powierzchnie")
			if nl=6 then
				set objPunkty=Application.CreateAppObject("Points")
			end if
		end if
		NowaLinia=false

		If Not Application.Map.AddFeature (objPunkt,false) Then
			MsgBox "Punktu nie dodano",vbExclamation,"Błąd"
			Pomiar=false
		End If		
	
		objPunkty.add objPunkt
		objPunkty.add objPunkt

end if
		
end sub

sub ObliczPowierzchnie

	if Powierzchnia and PowierzchniaTyp then
		'Wewnętrzna	
		
		if PowierzchniaWewNowa then
			'Jeśli nowy kontur zainicjuj wartości początkowe
			PPwewX=XC
			PPwewY=YC
			OPwewX=XC
			OPwewY=YC
			Owew=0
			PwewXmin=XC
			PowierzchniaWewNowa=false
		else
			'Licz obwód
			Owew=Owew-sqr((PPwewX-OPwewX)*(PPwewX-OPwewX)+(PPwewY-OPwewY)*(PPwewY-OPwewY))
			Owew=Owew+sqr((XC-OPwewX)*(XC-OPwewX)+(YC-OPwewY)*(YC-OPwewY))
			Owew=Owew+sqr((PPwewX-XC)*(PPwewX-XC)+(PPwewY-YC)*(PPwewY-YC))
			'Licz powierzcznię
			'Usuniecie ostatniego zamkniecia konturu
			if debugPow then msgbox Pwew,0,"1 Powierzchnia wewnętrzna"
			if debugPow then msgbox PowierzchniaTrapezu(OPwewX,OPwewY,PPwewX,PPwewY,PwewXmin),0,"2 Usuwamy zamknięcie"
			Pwew=Pwew-PowierzchniaTrapezu(OPwewX,OPwewY,PPwewX,PPwewY,PwewXmin)
			if debugPow then msgbox Pwew,0,"3 Powierzchnia bez zamkniecia"
			'Czy nie jesteśmy poniżej poziomu odniesienia
			if XC<PwewXmin then
				if debugPow then msgbox (PwewXmin-XC)*(OPwewY-PPwewY),0,"4 Korekta za zmianę poziomu odniesienia"
				if debugPow then msgbox XC,0,"5 Nowy poziom odniesienia"
				Pwew=Pwew+(PwewXmin-XC)*(OPwewY-PPwewY)
				if debugPow then msgbox Pwew,0,"6 Powierzchnia po korekcie"
				PwewXmin=XC
			end if
			'Dodanie nowego fragmentu
			if debugPow then msgbox PowierzchniaTrapezu(OPwewX,OPwewY,XC,YC,PwewXmin),0,"7 Dodajemy nowy element"
			Pwew=Pwew+PowierzchniaTrapezu(OPwewX,OPwewY,XC,YC,PwewXmin)
			OPwewX=XC
			OPwewY=YC
			'Zamkniecie konturu
			if debugPow then msgbox PowierzchniaTrapezu(OPwewX,OPwewY,PPwewX,PPwewY,PwewXmin),0,"8 Dodajemy nowe zamknięcie"
			Pwew=Pwew+PowierzchniaTrapezu(OPwewX,OPwewY,PPwewX,PPwewY,PwewXmin)
			if debugPow then msgbox Pwew,0,"Finalna powierzchnia wewnętrzna"
		end if

	end if 

	if Powierzchnia and not(PowierzchniaTyp) then
		'Zewnętrzna

		if PowierzchniaZewNowa then
			'Jeśli nowy kontur zainicjuj wartości początkowe
			PPzewX=XC
			PPzewY=YC
			OPzewX=XC
			OPzewY=YC
			Pzew=0
			Ozew=0
			PzewXmin=XC
			PowierzchniaZewNowa=false
		else
			'Licz obwód
			Ozew=Ozew-sqr((PPzewX-OPzewX)*(PPzewX-OPzewX)+(PPzewY-OPzewY)*(PPzewY-OPzewY))
			Ozew=Ozew+sqr((XC-OPzewX)*(XC-OPzewX)+(YC-OPzewY)*(YC-OPzewY))
			Ozew=Ozew+sqr((PPzewX-XC)*(PPzewX-XC)+(PPzewY-YC)*(PPzewY-YC))
			'Licz powierzcznię
			'Usuniecie ostatniego zamkniecia konturu
			if debugPow then msgbox Pzew,0,"1 Powierzchnia zewnętrzna"
			if debugPow then msgbox PowierzchniaTrapezu(OPzewX,OPzewY,PPzewX,PPzewY,PzewXmin),0,"2 Usuwamy zamknięcie"
			Pzew=Pzew-PowierzchniaTrapezu(OPzewX,OPzewY,PPzewX,PPzewY,PzewXmin)
			if debugPow then msgbox Pzew,0,"3 Powierzchnia bez zamkniecia"
			'Czy nie jesteśmy poniżej poziomu odniesienia
			if XC<PzewXmin then
				if debugPow then msgbox (PzewXmin-XC)*(OPzewY-PPzewY),0,"4 Korekta za zmianę poziomu odniesienia"
				if debugPow then msgbox XC,0,"5 Nowy poziom odniesienia"
				Pzew=Pzew+(PzewXmin-XC)*(OPzewY-PPzewY)
				if debugPow then msgbox Pzew,0,"6 Powierzchnia po korekcie"
				PzewXmin=XC
			end if
			'Dodanie nowego fragmentu
			if debugPow then msgbox PowierzchniaTrapezu(OPzewX,OPzewY,XC,YC,PzewXmin),0,"7 Dodajemy nowy element"
			Pzew=Pzew+PowierzchniaTrapezu(OPzewX,OPzewY,XC,YC,PzewXmin)
			OPzewX=XC
			OPZewY=YC
			'Zamkniecie konturu
			if debugPow then msgbox PowierzchniaTrapezu(OPzewX,OPzewY,PPzewX,PPzewY,PzewXmin),0,"8 Dodajemy nowe zamknięcie"
			Pzew=Pzew+PowierzchniaTrapezu(OPzewX,OPzewY,PPzewX,PPzewY,PzewXmin)
			if debugPow then msgbox Pzew,0,"Finalna powierzchnia zewnętrzna"
		end if
		
	end if 

end sub

function PowierzchniaTrapezu (x1,y1,x2,y2,xmin)
dim dx1,dx2,dy
	
	dx1=x1-xmin
	dx2=x2-xmin
	dy=y2-y1
	
	PowierzchniaTrapezu=0.5*(dx1+dx2)*dy
	
end function

Sub Inicjacja
dim x, f

	DebugGPS=false
	Set f = Application.CreateAppObject("file")
	if f.Exists ("\debugGPS.txt") then DebugGPS=true
		
	if DebugGPS then msgbox "DebugGPS mode"

	debugPow=false
	if DebugPow then msgbox "DebugArea mode"

	if DebugGPS then
		set objPlik=CreateAppObject("FILE")
		objPlik.open "\debugGPS.txt",apFileRead,apFileASCII
		x=objPlik.ReadLine
	end if

	'Interwał zliczania w sekundach dla odbiornika GPS
	GPSInterval=1

	pi=4*atn(1)
	'Real czas wykonania ostatniego pomiaru
	Czas=now()
	'Boolean Czy z GPS otrzymano nową pozycję
	NowaPozycja=False
	'Real ostatnie współrzędne z GPS
	X=0.0
	Y=0.0
	Z=0.0 
	'Real współrzędne ostatniego punktu z uwzględnieniem offsetu
	XC=0.0
	YC=0.0
	ZC=0.0
	'Real współrzędne poprzedniego punktu z GPS
	X0=0.0
	Y0=0.0
	Z0=0.0
	'Boolean Czy wyzwolić pomiar
	Pomiar=False
	'Boolean Czy pomiar ma być ciągły
	Ciagly=False
	'integer Typ offsetu 0-bez offsetu, 1 - azymut, 2 - domiar
	OffsetTyp=0
	'Boolean Czy pamiętać wartość offsetów
	OffsetPamietaj=false
	'Real OffsetTyp=1 Azymut offsetu [rad]
	Offset1Az=0.0
	'Real OffsetTyp=1 Odległość
	Offset1Dist=0.0
	'Real OffsetTyp=2 Odległość
	Offset2Dist=0.0
	'Boolean Czy ma być prowadzona aktualizacja wyświetlania wyników
	Aktualizacja=False
	'Boolean Czy punkt ma być uwzględniany w pomiarze powierzchni
	Powierzchnia=False
	'Boolean Czy powierzchnia wewnętrzna (true) czy zewnętrzna (false)
	PowierzchniaTyp=False
	PowierzchniaBlokada=False
	'Real Wartość powierzchni wewnętrznej
	Pwew=0.0
	'Real Wartość powierzchni zewnętrznej
	Pzew=0.0
	'Real Obwód powierzchni wewnętrznej
	Owew=0.0
	'Real Obwód powierzchni zewnętrznej
	Ozew=0.0
	'Pierwszy punkt powierzchni wewnętrznej
	PPwewX=0.0
	PPwewY=0.0
	'Ostatni punkt powierzchni wewnętrznej
	OPwewX=0.0
	OPwewY=0.0
	'Pierwszy punkt powierzchni zewnętrznej X
	PPzewX=0.0
	PPzewY=0.0
	'Ostatni punkt powierzchni zewnętrznej
	OPzewX=0.0
	OPzewY=0.0
	'Punkt odniesienia do pomiaru odległości
	PPX=0.0
	PPY=0.0

	

	set objFile=CreateAppObject("FILE")
	
	GPSnmeaTryb=0
	GPS_DM = false
	GPS_EM = false

	'Boolean Powierzcznia w hektarach
	wsPha=false
	'Boolean Powierzchnia w m2
	wsPm2=false
	'Boolean Obwód konturu zewnętrznego
	wsO=false
	'Boolean Powierzcznia wewnętrzna w hektarach
	wsPhaWew=false
	'Boolean Powierzchnia wewnętrzna w m2
	wsPm2Wew=false
	'Boolean Obwód konturu wewnętrznego
	wsOWew=false
	'Double Współrzędne wybranego punktu
	wsX=0.0
	wsY=0.0
	'Typ rysowanego obiektu
	wsPLB="Punkt"
	
	kPow=0.0
	kObw=0.0
	
	'Czy ropcocząć tysowanie nowej granicy
	NowaLinia=true
	'Tu są zapisane punkty załamania linii
	set objPunkty=Application.CreateAppObject("Points")
	
	Set objAPTimer = Application.Timer 
	objAPTimer.Interval =500
	objAPTimer.Enabled =True

End Sub


sub cmdPomiar
	objAPTimer.Enabled =false
	objAPTimer.Interval =1000
	objAPTimer.Enabled =true
	Pomiar=true
end sub

sub cmdPomiarStop
dim objLinia, objPunkt, objPlikShape
dim i, a

	Pomiar=false

if wsPLB="Linia" then	
	set objLinia=Application.CreateAppObject("Line")
	call objLinia.parts.add(objPunkty)
	call map.AddFeature(objLinia,false)
	set objPunkty=Application.CreateAppObject("Points")

	set objPunkt=Application.CreateAppObject("Point")
	objPunkt.X=XC
	objPunkt.Y=YC
	objPunkt.Z=ZC
	objPunkty.add objPunkt
end if

if wsPLB="Powierzchnia" then	

	set objLinia=Application.CreateAppObject("Polygon")
	call objLinia.parts.add(objPunkty)
	call map.AddFeature(objLinia,True)
	set objPunkty=Application.CreateAppObject("Points")

	set objPunkt=Application.CreateAppObject("Point")
	objPunkt.X=XC
	objPunkt.Y=YC
	objPunkt.Z=ZC
	objPunkty.add objPunkt
	objPunkty.add objPunkt
	
	call test(1)

	dim sl, shtype, sh, rs
	set sl=Application.Map.SelectionLayer
	if sl is nothing then exit sub
	
	set rs=sl.Records
	set sh=rs.Fields.Shape
	shtype=sh.ShapeType
	if shtype=5 then 
	rs.Fields("Pow0").value=cstr(int(tempPow))
	rs.Fields("Obw0").value=cstr(int(tempObw))
	rs.Update
	end if
end if
PowierzchniaBlokada=false
NowaLinia=true


	set sl=Application.Map.SelectionLayer
	if sl is nothing then exit sub
	
	set rs=sl.Records
		
	rs.Fields("Producent").value=cstr(Wlasciciel)
	rs.Fields("Inspektor").value=cstr(Inspektor)
	rs.Fields("NrProd").value=cstr(int(NrProducenta))
	rs.Fields("Data pom").value=cstr(Data)
	rs.Update

end sub

sub KoniecPomiaruKonturu

	Powierzchnia=not(Powierzchnia)
	if powierzchnia then
		msgbox "Wstrzymanie pomiaru powierzchni",vbOkOnly,"Powierzchnia"
		PowierzchniaBlokada=true
	else
		msgbox "Kontynuacja pomiaru powierzchni",vbOkOnly,"Powierzchnia"
		PowierzchniaBlokada=false
	end if

end sub

sub KonfiguracjaGPSOnLoad(ppage)
dim g_pKonfPageControls
	set g_pKonfPageControls=ThisEvent.Object.Pages(1).Controls
	set g_pKonfPageControls("Combo1").value=wsPRej
	set g_pKonfPageControls("Combo2").value=wsPLB
	set g_pKonfPageControls("Combo3").value=GPSInterval
	
end sub

sub cmdParPomiaru
dim ppage, GPSInt
	
	set ppage=ThisEvent.Object.Pages(1).Controls

		if 1=ppage("Combo1").value then
			Ciagly=True
		else
			Ciagly=False
		
	end if
	wsRej=ppage("Combo1").value
	wsPLB=ppage("Combo2").value

	If wsPLB ="Linia" then
		NowaLinia=True
	end if

	If wsPLB ="Powierzchnia" then
		NowaLinia=True
	end if

	GPSInt=ppage("Combo3").value
	if isnumeric(GPSInt) then
		GPSInterval=int(GPSInt)
	else
		GPSInterval=1
	end if

	applet.forms("Form2").close

debug=false
If debug then
	if ciagly then
		msgbox "Ciagly",vbokonly,"Pomiar"
	else
		msgbox "Pojedynczy",vbokonly,"Pomiar"
	end if
	
end if

end sub

sub autor
msgbox "Autor:"&vbcr&vblf& _
"Łukasz Wiszniewski"&vbcr&vblf& _
"Copyright 2010"&vbcr&vblf& _
"Version 1.0.8",vbokonly,"O autorze"
end sub
