Option Explicit
Dim con,rs

Set con=createobject("adodb.connection")
Set rs=createobject("adodb.recordset")

con.open"Driver={MySQL ODBC 5.7 ANSI Driver};Server=localhost;Database=sakila;User=root;Password=Passw0rd!"
rs.open "select * from actor",con


Do while not rs.eof
	VbWindow("Form1").VbEdit("val1").Set rs.fields("v1")
	VbWindow("Form1").VbEdit("val2").Set rs.fields("v2")
	VbWindow("Form1").VbButton("ADD").Click
	rs.movenext
Loop

'Release objects
Set rs= nothing
Set con= nothing

