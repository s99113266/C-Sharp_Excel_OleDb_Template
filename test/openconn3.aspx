<%
'''ASP方法开启资料库必须在page宣告aspcompat="true"(用在新曾、修改、删除)
dim usdb = Request.PhysicalApplicationPath & "db/Database.accdb"
dim con
if My.computer.FileSystem.FileExists(usdb) then
    con = server.CreateObject("ADODB.Connection")
    con.open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & usdb & ";Jet OLEDB:Database Password=tn999kinggnik999nt")
end if
%>