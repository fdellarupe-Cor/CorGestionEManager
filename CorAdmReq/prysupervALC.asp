<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))



  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar



  sopcion = request.querystring("respon")





    sImagS = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=3>&#0252;</FONT></B>"
    sImagN = "<B><FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=4>&#0251;</FONT></B>"

    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenA = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"

    sImageR = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>=</FONT></B>"
    sImageS = "<B><FONT FACE=WEBDINGS COLOR=#EEEEEE SIZE=3>4</FONT></B>"
    sImageT = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>4</FONT></B>"
    sImageG = "<B><FONT FACE=WEBDINGS COLOR=#007700 SIZE=3>&#0252;</FONT></B>"


  StrSql = "Select * From PROYRESP where id = 0" & sopcion
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","TTT colspan=2","","","","","",sImagen & "&nbsp;PERMISOS EJECUTIVOS","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

  xTrColor = "#FFFFFF" 

  if not RS.EOF then




  celdas = ofv.GenCelda("","TT ","center","20%","","","","USUARIO:","si")


    if RS(7) = "NO" then
       sInput = "<B><I>SIN AFECTAR</I></B>"

    else

      strsql = "Select * From USUARIOS where ID = '" & rs(7) & "' "
      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsniv.eof then sniv = rsniv(1)
      call ofv.cerrarconsulta(rsniv)
      sInput = "<B>" & sniv & "</B>"

    end if

    celdas = celdas & ofv.GenCelda("","VT","","80%","","","",sInput,"si")



  filas = filas & ofv.GenRow("","","","","22","","",celdas,"si")

  celdas = ofv.GenCelda("","TTR colspan=2 ","center","","","","","ALCANCE","si")


  filas = filas & ofv.GenRow("","","","","22","","",celdas,"si")


       sInput = "<FONT COLOR=#993333 >GLOBAL</FONT>"
       sInput = "<A CLASS=NT HREF='PRYSUPERV.ASP?OPCION=RS&ID=" & RS(0) & "&SECTOR=999999' >" & SINPUT & "</A>"



                         if (RS(6) & "") = "999999" then
                            xTrColor = "#F9F9F9" 
                            sImage   = sImageT
                         ELSE
                            xTrColor = "#FFFFFF" 
                            sImage   = sImageS
                         END IF




  celdas = ofv.GenCelda("","UTI colspan=2 ","","","","","",sImage & "&nbsp;" & sInput,"si")


  filas = filas & ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","22","","",celdas,"si")


    sInput = "---"

    if (rs(4) & "") = "" then
       sInput = "---"

    elseif (rs(4) & "") = "999999" then

       sInput = "<FONT COLOR=#993333 >GLOBAL</FONT>"

    else

     StrSql = "Select * From ECO where ECO = 0" & RS(4)
     set rs2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rs2.eof then
        sInput = rs2(1)
        sInput = "<FONT COLOR=#6699CC >" & sInput & "</FONT>"
        sInput = "<A CLASS=NT HREF='PRYSUPERV.ASP?OPCION=RS&ID=" & RS(0) & "&SECTOR=" & RS(4) & "' >" & SINPUT & "</A>"
     ELSE
        sInput = "NO DETERMINADO"
        sInput = "<FONT COLOR=#6699CC >" & sInput & "</FONT>"
     END IF


    end if


                         if (RS(6) & "") = (CSTR(RS(4)) & "") then
                            xTrColor = "#F9F9F9" 
                            sImage   = sImageT
                         ELSE
                            xTrColor = "#FFFFFF" 
                            sImage   = sImageS
                         END IF




  celdas = ofv.GenCelda("","UTI colspan=2 ","","","","","",sImage & "&nbsp;" & sInput,"si")


  filas = filas & ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","22","","",celdas,"si")


    sInput = "---"

    if (rs(3) & "") = "" then
       sInput = "---"

    elseif (rs(3) & "") = "999999" then

       sInput = "<FONT COLOR=#993333 >GLOBAL</FONT>"

    else

     StrSql = "Select * From ECO where ECO = 0" & RS(3)
     set rs2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rs2.eof then
        sInput = rs2(1)
        sInput = "<FONT COLOR=#008800 >" & sInput & "</FONT>"
        sInput = "<A CLASS=NT HREF='PRYSUPERV.ASP?OPCION=RS&ID=" & RS(0) & "&SECTOR=" & RS(3) & "' >" & SINPUT & "</A>"
     ELSE
        sInput = "NO DETERMINADO"
        sInput = "<FONT COLOR=#008800 >" & sInput & "</FONT>"
     END IF


    end if

                         if (RS(6) & "") = (CSTR(RS(3)) & "") then
                            xTrColor = "#F9F9F9" 
                            sImage   = sImageT
                         ELSE
                            xTrColor = "#FFFFFF" 
                            sImage   = sImageS
                         END IF



  celdas = ofv.GenCelda("","UTI colspan=2 ","","","","","",sImage & "&nbsp;" & sInput,"si")


  filas = filas & ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","22","","",celdas,"si")










   end if






    sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"



  'Botonera



  sAccion = "PRYSUPERV.asp"
  VOLVER = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","",VOLVER)


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML








%>