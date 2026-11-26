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

  session("NCliente") = "0"

  sopcion = request.querystring("opcion")

  if sopcion = "SI" THEN
     sECTOR  = request.form("SECTOR")
     susrafc = request.form("USRAFC")
     XXX     = GENSECTORSUP(sECTOR,susrafc)
  END IF

  if sopcion = "RS" THEN
     sECTOR  = request.querystring("SECTOR")
     susrafc = request.querystring("ID")
    '' RESPONSE.WRITE sECTOR & " " & susrafc
     StrSql  = "UPDATE PROYRESP SET SUBTIPO = '" & sECTOR & "' WHERE ID = 0" & susrafc
     CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
  END IF



 XSEL1 = " and skr <> 'A' "
 XSEL2 = " and skr = 'A' "

 if sopcion = "RF" THEN
    XXX=GENREFRESH()
 END IF

    sImagS = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=3>&#0252;</FONT></B>"
    sImagN = "<B><FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=4>&#0251;</FONT></B>"

    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenA = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"

    sImageR = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>=</FONT></B>"
    sImageS = "<B><FONT FACE=WEBDINGS COLOR=#6699CC SIZE=3>4</FONT></B>"
    sImageG = "<B><FONT FACE=WEBDINGS COLOR=#007700 SIZE=3>&#0252;</FONT></B>"


  StrSql = "Select * From PROYRESP where codigo = '" & session("NCliente") & "' ORDER BY ORDEN "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","TTT colspan=8","","","","","",sImagen & "&nbsp;PERMISOS EJECUTIVOS","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","TTR COLSPAN=4","center","","","","","USUARIOS","si")
'  celdas = celdas & ofv.GenCelda("","TTR","center","","","","","---","si")
  celdas = celdas & ofv.GenCelda("","TTR colspan=2 ","center","","","","","ALCANCE","si")
  celdas = celdas & ofv.GenCelda("","TTR","center","","","","","CONTROL","si")
  celdas = celdas & ofv.GenCelda("","TTR","center","","","","","AVANCE","si")

  filas = filas & ofv.GenRow("","","","","22","","",celdas,"si")

  xTrColor = "#FFFFFF" 

  DO UNTIL RS.EOF


    'Eliminar
    if bSuprimir AND RS(5) = "AF" then




         sHref = "asp/EliminarRegistro.asp?ID=" & rs(0)
         sHref = sHref & "&Tabla=PROYRESP&volver=" & vuelta & "&Eliminar=NSNC"
         CeldaS = ofv.BotonEliminar(" RESPONSABLE ",sHref)
         CeldaS = REPLACE(CeldaS,"<TD","<TD WIDTH=1%  ")
    ELSE
         CeldaS = ofv.GenCelda("","UT","","1%","","","","&nbsp;","si")
    END IF



    if RS(5) = "OR" then
       sInput = sImagenO
    ELSEif RS(5) = "AF" then
       sInput = sImagenA
    ELSE
       sInput = sImagenI
    END IF

    celdas = celdas & ofv.GenCelda("","VT","","1%","","","",sInput,"si")


    if RS(7) = "NO" then
       sInput = "<B><I>SIN AFECTAR</I></B>"

    else

      strsql = "Select * From USUARIOS where ID = '" & rs(7) & "' "
      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsniv.eof then sniv = rsniv(1)
      call ofv.cerrarconsulta(rsniv)
      sInput = "<B>" & sniv & "</B>"

    end if

    celdas = celdas & ofv.GenCelda("","VT","","40%","","","",sInput,"si")

    SFORM = "Usuarios" & RS(0)

    SFUNC = RS(15) & ""

    if SFUNC = "" then
       sInput = "<B><I>---</I></B>"

    else

      strsql = "Select * From PROYFUNCIONES where ID = " & rs(15)
      set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if not rsniv.eof then sniv = rsniv(1)
      call ofv.cerrarconsulta(rsniv)
      sInput = "<B><I>" & sniv & "</I></B>"

    end if


       celdas = celdas & ofv.GenCelda("","VT ","","5%","","","",sInput,"si")

  '  sInput = "---"
  '  celdas = celdas & ofv.GenCelda("","VT","CENTER","10%","","","",sInput,"si")

    sInput = "---"

    if (rs(6) & "") = "" then
       sInput = "---"

    elseif (rs(6) & "") = "999999" then

       sInput = "<A CLASS=NT HREF='PRYSUPERVALC.ASP?RESPON=" & RS(0) & "' ><FONT COLOR=#993333 >GLOBAL</FONT></A>"

    else

     StrSql = "Select * From ECO where ECO = 0" & RS(6)
     set rs2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rs2.eof then
        sInput = rs2(1)
        sInput = "<FONT COLOR=#008800 >" & sInput & "</FONT>"
        sInput = "<A CLASS=NT HREF='PRYSUPERVALC.ASP?RESPON=" & RS(0) & "' >" & SINPUT & "</A>"

     ELSE
        sInput = "NO DETERMINADO"
        sInput = "<FONT COLOR=#008800 >" & sInput & "</FONT>"
     END IF




    end if



    celdas = celdas & ofv.GenCelda("","uT colspan=2","","25%","","","",sInput,"si")

    scomK = "1" & chr(9) & "0"
    scomD = "SI" & chr(9) & "NO"

    sInput = ofv.gencomboFX(sForm,rs(11).name,rs(11),scomK,scomD," ")

 '   sInput = ofv.combosino(sForm,rs(11),rs(11).name)
     IF RS(11) = 1 THEN
   '    SINPUT = SINPUT & SIMAGS
     ELSE
   '   SINPUT = SINPUT & SIMAGN
     END IF 
    celdas = celdas & ofv.GenCelda("","VT","CENTER","5%","","","",sInput,"si")

    sInput = ofv.gencomboFX(sForm,rs(12).name,rs(12),scomK,scomD," ")

  '  sInput = ofv.combosino(sForm,rs(12),rs(12).name)
     IF RS(12) = 1 THEN
  '    SINPUT = SINPUT & SIMAGS
     ELSE
  '     SINPUT = SINPUT & SIMAGN
     END IF 
    celdas = celdas & ofv.GenCelda("","VT","CENTER","5%","","","",sInput,"si")



    if bModificar then
      sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=PROYRESP"
      formulario0 = ofv.GenForm(SFORM,"",sAccion,"","post","",celdas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",celdas,"si")
    end if


    filas = filas & ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","10","","",formulario0,"si")



                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF


    RS.MOVENEXT
    LOOP





    Celda = ofv.GenCelda("","abc colspan=8 style='border-bottom-style:RIDGE;border-bottom-width:2;border-bottom-color:#F7AA22;FONT-SIZE:5pt' ","center","100%","","","","&nbsp;","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")


    sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"



  'Botonera



 if bAgregar and sopcion <> "AT" THEN

    sAccion = "prysupervBUS.asp?VASP=ES"
    Agregar = ofv.BotonAgregar(sAccion)

 ELSE
    Agregar = ""
 END IF


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","",Agregar)


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML








%>