<!-- #include virtual="/NetPower2.asp" --><!-- #include file="../Common.asp" -->
<!-- #include virtual="/Server/cblib.asp" -->
<!-- #include virtual="/Server/public.asp" -->
<% 
dim ID, rs

Sub Load()
	If IsEmpty(ID) then ID = ClngEx(stdin("ID"))
	Set rs = connEx.execute("select * from FindPet where id=" & ID)
	if rs.eof then Err.Raise 10000, ModName & ".Load", "NotExistsID"
	'if Clng(rs("IsDeleted").value) = 1 then Err.Raise 10000, ModName & ".Load", "Deleted"
	fnMapping
	np2.mapBind stdout
	fnLoad()
	LogWrite EventName & ".´ò¿ª", "ID=" & ID & ",  Title=" & rs("placeText").value
	set stdin = Nothing
End Sub

Function fnMapping()
end Function

Function fnLoad()
end Function

Load
Function fnMapping()
	map ModName, "AddNew","",MAPVOID,"cbRSAddNew"
	map ModName,"ID","ID",MAPINT,"cbRSDirect"
	map ModName,"isRe","isRe",MAPINT,"cbRSDirect"
	map ModName,"isdeleted","isdeleted",MAPINT,"cbRSDirect"
	map ModName,"timeInput","timeInput",MAPSTRING,"cbRSDirect"
	map ModName,"placeText","placeText",MAPSTRING,"cbRSDirect"
	map ModName,"placepoint","placepoint",MAPSTRING,"cbRSDirect"
	map ModName,"Varieties","Varieties",MAPSTRING,"cbRSDirect"
	map ModName,"Gender","Gender",MAPSTRING,"cbRSDirect"
	map ModName,"sterilization","sterilization",MAPSTRING,"cbRSDirect"
	map ModName,"color","color",MAPSTRING,"cbRSDirect"
	map ModName,"isname","isname",MAPSTRING,"cbRSDirect"
	map ModName,"phone","phone",MAPSTRING,"cbRSDirect"
	map ModName,"memo","memo",MAPSTRING,"cbRSDirect"
	map ModName,"imgText","imgText",MAPSTRING,"cbRSDirect"
	map ModName,"state","state",MAPSTRING,"cbRSDirect"
	map ModName,"resulttext","resulttext",MAPSTRING,"cbRSDirect"
End Function

Function fnLoad()
	np2.mapping 0, stdout, ModName, rs
end Function

%>