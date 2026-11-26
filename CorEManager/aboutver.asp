<%
Public Conn
public  strconn
strconn1="dsn=DSNCOR01ESSsql"


function conectar (strconn)

  set conn = server.createobject("adodb.connection")
  conn.open strconn
  set conectar=conn
end function


function CrearConsultaCn(strsql)
dim rs
  set rs=conn.execute(strsql)
  set crearconsultaCn=rs

end function

function CrearConsulta(strsql,strconn)

dim rs
 set rs=server.createobject("ador.recordset")

 rs.open strsql,strconn
 set crearconsulta=rs

end function

function CrearConsultaEx(byref strsql,byref conn,TipoConsulta, byref a)
dim t
dim rs
dim errloop
dim errores
 set rs=server.createobject("adodb.command")



  rs.activeconnection = conn
  rs.commandtext = strsql
  rs.commandtype = TipoConsulta


select case TipoConsulta
 Case 1
   on error resume next
   set crearconsultaEx= rs.execute()
   set errores=conn.errors
   if errores.count > 0 then
'     response.write "<script> window.open('../asp/errordb.asp?volver=../PerfilesdeEsquemaNormativo.asp', 'displayWindow','width=255,height=100,status=no,toolbar=no,menubar=no');</script>"
   '   response.redirect "../asp/errordb.asp?volver="& a
     '  response.write errores(1).description
   end if
 Case 4
   convertirArregloParam parametros,carblanco," "
   set crearconsultaEx= rs.execute( ,a)
End Select
end function

sub cerrarconsulta(byref consulta)
   consulta.close
   set consulta = nothing
end sub

function ConsultaVacia(Byref RsRecordset)
  ConsultaVacia =(RsRecordset.Eof and RsRecordset.bof)
end function

sub cerrarconn(byref conn)
 conn.close
 set conn = nothing
end sub

response.write "<html><body><p><font face=tahoma size=1 style={color:#336699;}>"
  Dim fso,Folder,Files,F, s, n

  set cn = conectar(strconn1)
  Strsql = "Select * from CorConsulta"
  set rs = crearconsultaEx(StrSql,cn,1,parametros)
  if Not rs.EOF then
     xLicencia = rs(2)
     xEmpresa = rs(3)
  end if

  call cerrarconsulta(rs)
  call cerrarconn(cn)

  response.write "Cor ESolution Suite V4.6 <br>"
  response.write "Licencia " & xLicencia & " otorgada a " & xEmpresa & "<br>"

  Set fso = CreateObject("Scripting.FileSystemObject")
  sPath = server.mappath("../coressdll")
  response.write spath & "<br>"
  Set Folder = fso.getfolder(sPath)
  Set Files = Folder.Files


  For Each F in Files
     xVersion = ""
     if len(f.type) > 3 then
        if ucase(left(f.type,3)) = "APP" then
           Set ver = nothing
           on error resume next
           xnombre = left(f.name,instr(Ucase(f.name),".DLL") - 1)
           Set ver = CreateObject(xnombre & ".about")
           xVersion = ver.version
           if xVersion <> "" then
           response.write "Componente: " & xVersion & " " & "(" & f.name & ")" & " " & f.datecreated & "<br>"
           end if
        end if
     end if
  Next
response.write "</font></p></body></html>"

%>