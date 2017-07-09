<%
dim connEx

getConnEx()

sub getConnEx()
	set connEx = Server.CreateObject("ADODB.Connection")   
    connEx.Open "provider=sqloledb;data source=.;User ID=sa;pwd=1qaz!QAZ;DATABASE=FindPet;"   
end sub

%>