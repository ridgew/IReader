<%@language="vbscript" codepage="936"%>
<%
Option Explicit
On Error Resume Next

Const CountBookRead = True
Const EXPIREDDAYNUM = 500

Const BookRootPath = "D:\wwwroot\JackWind\wwwroot\wwwroot\#books\"


'Response.Write(Server.MapPath("/"))
'Response.End()

Const ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;user id=admin;Jet OLEDB:Database Password=VISNWORKS;Data Source=D:\wwwroot\JackWind\databases\iReaderBook2005.mdb;"
'Const ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;user id=admin;Jet OLEDB:Database Password=VISNWORKS;Data Source=E:\LocalDev\wwwroot\iReader\bookDb.mdb;"

%>