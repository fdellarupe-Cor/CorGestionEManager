<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  bModificar = False
  ofv.ObtenerAtributos Session("Form"),""
  bModificar = ofv.AModificar

 set cn = ofv.conectar(ofv.strconn6)

    sImagenP = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>&#0143;</FONT></B>"
    sImagenR = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>2</FONT></B>"  


'xxxxxxxxxxxxxxxxxx
  StrSql = "Select * From REQUERIMIENTO WHERE CODIGO = '" & SESSION("NCLIENTE") & "' "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
        TITULO = rs(2)
        ASIGNA = rs(14)
        treq   = rs(15)
        cinter = rs(19) & ""
     end if

dim atab(15,1)
dim atabN(15,1)


    for xx = 1 to 15
           atab(xx,0) = 99
           atabN(xx,0) = 99
    next

    for xx = 1 to 9

           atab(XX,0) = xx + 2



    next



     StrSql = "Select * From TIPOREQ WHERE TIPOREQ = '" & treq & "' "
     set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rsX.eof then




        for yy = 1 to 15
 
            if clng(atab(yy,0)) = 99 then exit for
               xx = atab(yy,0)

               IF RSX(XX) THEN
                  atabN(yy,0) = XX
                  atabN(yy,1) = rsX(xx + 15)
               ELSE
                  atabN(yy,0) = 99
               END IF 
        next

     end if



  sHTML = "<html>"

     wxs = "<script language=JavaScript>" & Chr(13)
     wxs = wxs & " function IniWin() { " & Chr(13)
     wxs = wxs & " document.ICONACT.submit(); " & Chr(13)
     wxs = wxs & " } " & Chr(13)
     wxs = wxs & "function AbrirWin() { " & Chr(13)
     wxs = wxs & " document.festad.submit(); " & Chr(13)
     wxs = wxs & " } " & Chr(13)
     wxs = wxs & "function closWin() { " & Chr(13)
     wxs = wxs & " document.init.submit(); " & Chr(13)
     wxs = wxs & " } " & Chr(13)
     wxs = wxs & "</script>" & Chr(13)

    sHTML = sHTML & wxs



     sHTML = sHTML & "<body  onload='IniWin();' onunload='AbrirWin();' >"




  bstitulo = " " & TITULO

  xsCorEsq = ucase(bstitulo)


    IF  Session("FULLGUI") = "N" THEN
        styl = "background-color:#888888;"
    else     
        styl = XspEx("#888888","#cccccc",0) & ";"
    end if

    xsclase = "at"

    sHTML = sHTML & OFV.GenTabla("DBpanel", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "2", "", "", "", "no")

    h1 = OFV.gencelda("", "at style='padding-left:3px;' ", "", "2", "", "", "", "<IMG valign=middle border=0 src='../coress20/imagenes/corlogo.gif' style='filter: chroma(color=#ffffff);' >", "si")
    h1 = h1 & ofv.GenCelda("",xsclase,"","","","","",xsCorEsq,"si")

    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
    sHTML = sHTML & ofv.GenFTabla()




    xverasp = "BarraPRYEnCierra2.asp?"
    xverasp = xverasp & "USR=" & REPLACE(Session("Usuario") & "*"," ","_")
    xverasp = xverasp & "&NETUSR=" & REPLACE(Session("UsuarNT") & "*"," ","_")
    xverasp = xverasp & "&DMN=" & REPLACE(Session("DOMINIO") & "*"," ","_")
    xverasp = xverasp & "&LID=" & REPLACE(SESSION("LOGIDUSRDET") & "*"," ","_")
    xverasp = xverasp & "&LIDOBJ=" & REPLACE(SESSION("LOGIDDET") & "*"," ","_")

    sHTML = sHTML & "<form name=festad action='" & xverasp & "' " 
    sHTML = sHTML & " method=post target=Estad style='display:none;' >"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"

    xSolasp = ""  
    xSolasp = xSolasp & "BarraREQIcon.asp?tsk=" & tsk & "&SSCLICK=99"

    sHTML = sHTML & "<form name=ICONACT action='" & xSolasp  & "' "
    sHTML = sHTML & " method=post target='WK2' style='display:none;' >"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"

     wxs = "<script language=JavaScript>" & Chr(13)
     wxs = wxs & " window.top.CorEsq = '" & xsCorEsq & "'; " & Chr(13)
     wxs = wxs & " window.top.focus(); " & Chr(13)
     wxs = wxs & "</script>" & Chr(13)

    sHTML = sHTML & wxs



  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"


end if

Response.Write sHTML

%>
