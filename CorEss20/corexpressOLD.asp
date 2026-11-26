<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorInit.inc" -->
<!-- #INCLUDE FILE="IncFile/Corval.inc" -->
<%

' http://server-a/coress28/coress20/corexpress.asp?opt=ED&id=103-MNN-1&durl=codigo=103_tipo=MN_perfil=OBJARC
function editen(sElemento)

'sElemento = request.querystring("ID")
n = instr(sElemento,"-")
nCodobjeto = left(sElemento,n-1)
n2 = instr(n+1,sElemento,"-")
sTipobjeto = mid(sElemento,n+1,n2-(n+1))
sEstado = mid(sElemento,n2+1)
sEsqSup = trim(sTipobjeto)
sEsqSup = mid(sEsqSup,len(sEsqSup))
sTipobjeto = left(sTipobjeto,len(sTipobjeto)-1)

ofv.ObtenerAtributos Session("Form2"),""
bAgregar = ofv.AAgregar
bModificar = ofv.AModificar
bSuprimir = ofv.AEliminar
bConsultar = ofv.AConsultar

set cn = ofv.conectar(ofv.strconn2)

scodigo = nCodobjeto
stipo = sTipobjeto

srestric = "N"



'xxxxxxxxxxxxxxxxxxxxxxxx

if sEsqSup <> "Y" then
    redim preserve sPath(0)
    i = 1
    do

       strsql = "select * from csbarr where codojasoc = " & sCodigo
       strsql = strsql & " and tipobjasoc = '" & sTipo & "'"
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

      if not rs.eof then
        redim preserve sPath(i)
        sCodigo = rs(1)
        sTipo = rs(2)
        if not isnull(isnull(rs(0))) then
          sPath(i) = rs(0)
        else
          i = i + 1
          exit do
        end if
        i = i + 1
      else
        exit do
      end if
      call ofv.cerrarconsulta(rs)
    loop

end if

    strsql = "SELECT * FROM PRODSERV where codprodserv = " & sCodigo
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs.eof then
       if rs("segaceso") <> 0 or rs("segactual") <> 0 then srestric = "S"
    end if
    call ofv.cerrarconsulta(rs)

    if srestric = "S" then
            StrSql = "Select * From subtemaskr "
            StrSql = StrSql & "Where codprodserv = " & sCodigo & " and tiposeg = 'T' "
            StrSql = StrSql & " and (grpusr = 'USR-" & Ucase(session("usuario"))
            StrSql = StrSql & "' or grpusr = 'GRP-" & Ucase(session("seggrp")) & "')"
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rs.eof then
          srestric = "N"
       else 
          srestric = "S"
       end if
       call ofv.cerrarconsulta(rs)
    end if



if srestric = "S" then
   bAgregar = 0
   bModificar = 0
   bSuprimir = 0
end if


'xxxxxxxxxxxxxxxxxxxxxxxx




session("Nomodificar") = "N"
if isnumeric(sEstado) then
if Session("Form2") <> "TEMAS" And sEstado = 3 then
  bSuprimir = False
  session("Nomodificar") = "S"
end if
if Session("Form2") <> "DESARROLLO" And sEstado = 1 then
  bSuprimir = False
  session("Nomodificar") = "S"
end if
if Session("Form2") <> "PARA APROBAR" And sEstado = 2 then
  bSuprimir = False
  session("Nomodificar") = "S"
end if
end if

if Session("Form2") = "TEMAS" or Session("Form2") = "ELIMINADAS" then
  bAgregar = False
end if

session("EsqSup") = sEsqSup



Session("Segct") = "S"
Session("Segmo") = "S"
Session("Segco") = "S"
Session("Segce") = "S"

if sEsqSup = "Y" then
  Strsql = "SELECT SUBTEMA.* FROM PRODSERV, SUBTEMA where "
  Strsql = Strsql & "PRODSERV.marcaps = SUBTEMA.tipobjeto and "
  Strsql = Strsql & " prodserv.marcaps = '" & sTipobjeto & "'"
else
  Strsql = "Select * From Tipobjeto Where tipobjeto='" & sTipobjeto & "'"
end if

set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

x = 0
redim Niveles(1)
if not rs.eof then
  sNombre = rs(1)
  sPerfil = rs(3)
  for z = 5 to 14  
    if rs(z) Then
      x = x + 1
      redim preserve Niveles(x)
      Niveles(x) = right(rs(z).name,1)
    end if
  next
end if

call ofv.cerrarconsulta(rs)



if sEsqSup = "N" then

   set cn = ofv.conectar(ofv.strconn2)

   Strsql = "Select * From OBJSDK Where tipobjeto='" & sTipobjeto & "' and codobjeto = "  & nCodobjeto
   set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if not rs.eof then

      
      if Session("Usuario") <> rs(7) then

         if rs(3) then Session("Segct") = "N"
         if rs(4) then Session("Segmo") = "N"
         if rs(5) then Session("Segco") = "N"
         if rs(6) then Session("Segce") = "N"

         call ofv.cerrarconsulta(rs)
         Strsql = "Select * From OBJSDKDET Where tipobjeto='" & sTipobjeto
         Strsql = Strsql & "' and codobjeto = "  & nCodobjeto
         Strsql = Strsql & " and ((grpusr = 'USR-" & Session("Usuario") & "') " 
         Strsql = Strsql & " or (grpusr = 'GRP-" & Session("Seggrp") & "')) " 
         set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         do until rs.eof 
            
            select case rs(3)
                   case "1"
                        Session("Segct") = "S"
                        Session("Segmo") = "S"
                        Session("Segco") = "S"
                        Session("Segce") = "S"

                   case "2"
                        Session("Segmo") = "S"
                        Session("Segco") = "S"

                   case "3"
                        Session("Segco") = "S"

                   case "4"
                        Session("Segco") = "S"
                        Session("Segce") = "S"

            end select
         rs.movenext
         loop
      end if
    end if
    call ofv.cerrarconsulta(rs)

    if srestric = "S" then
       Session("Segct") = "N"
       Session("Segmo") = "N"
       Session("Segce") = "N"
    end if


    if Session("Segce") = "S" and session("Nomodificar") <> "S" then
       session("Nomodificarce") = "N"
    else
       session("Nomodificarce") = "S"
    end if

    if Session("Segct") <> "S" and Session("Segmo") <> "S" then
       session("Nomodificar") = "S"
    end if



 ' aspeditar = "Editar.asp?Codigo=" & nCodobjeto & "&Tipo=" & sTipobjeto & "&Perfil=" & sPerfil
 
 ' aspeditar = cstr(aspeditar)
'  response.write aspeditar

 ' if bConsultar and Session("Segco") = "S" then sHTML = sHTML & ofv.editar(cstr(aspeditar))



end if


call ofv.cerrarconn(cn)



end function


  set oIni = server.createobject("CorSecurityTCdc.Inicio")

  Set oIni.Request = Request
  Set oIni.Server = Server
  Set oIni.Response = Response
  Set oIni.Session = Session
  Set oIni.Ofv = Ofv

  UserNt = UCase(Request.ServerVariables("LOGON_USER"))
  UserNt = Mid(UserNt, InStrRev(UserNt, "\") + 1, Len(UserNt))

  Session("Usuario") = UserNt

  session("noved") = "NO"



   xid = Request.querystring("id")
   xopt = Request.querystring("opt")
   xacceso = Request.querystring("durl")
   xacceso = replace(xacceso,"_","&")



    sHTML = Valida

    if len(shtml) > 1 then
       Response.write sHTML
    else
       If sHTML = "1" or sHTML = "2"  or sHTML = "3"  or sHTML = "6"  or sHTML = "7"  or sHTML = "8"  or sHTML = "9" Then
          sHTML = oIni.Inicio(sHTML)
          Response.write sHTML
       ElseIf sHTML = "0" Then
          xcinit = UsrInit()

          if xopt = "CC" then
             Response.redirect "../Coremanager/Consultar.asp?" & xacceso & "&Datos=*smenu=PORTEXTO"
          else
             if xopt = "EV" then Session("Form2") = "TEMAS"
             if xopt = "ED" then Session("Form2") = "DESARROLLO"
             if xopt = "EA" then Session("Form2") = "PARA APROBAR"

             editen(xid)
             Response.redirect "../Coremanager/editar.asp?" & xacceso & "&Datos=*smenu=PORTEXTO"
          end if
       Else
          Response.write sHTML
       End If
    end if 


%>
