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



 sopcion = request.querystring("opcion")

 if sopcion = "SI" THEN
    sECTOR = request.querystring("SECTOR")
    XXX=GENSECTORAF(sECTOR,"NO")
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


  StrSql = "Select * From PROYRESP where codigo = '" & session("NCliente")  & "' " & XSEL1 & " ORDER BY ORDEN "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","TTT colspan=7","","","","","",sImagen & "&nbsp;PERMISOS EJECUTIVOS","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","TTR COLSPAN=3","center","","","","","USUARIOS","si")
  celdas = celdas & ofv.GenCelda("","TTR","center","","","","","ACCESO","si")
  celdas = celdas & ofv.GenCelda("","TTR","center","","","","","CONTROL","si")
  celdas = celdas & ofv.GenCelda("","TTR","center","","","","","AVANCE","si")
  celdas = celdas & ofv.GenCelda("","TTR","center","","","","","CARPETA","si")
  filas = filas & ofv.GenRow("","","","","22","","",celdas,"si")



  DO UNTIL RS.EOF

    if RS(5) = "OR" then
       sInput = sImagenO
    ELSEif RS(5) = "AF" then
       sInput = sImagenA
    ELSE
       sInput = sImagenI
    END IF

    celdas = ofv.GenCelda("","VT ","","1%","","","",sInput,"si")


    if RS(7) = "NO" then
       sInput = "<B><I>SIN AFECTAR</I></B>"

    else

      strsql = "Select * From USUARIOS where ID = '" & rs(7) & "' "
      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsniv.eof then sniv = rsniv(1)
      call ofv.cerrarconsulta(rsniv)
      sInput = "<B>" & sniv & "</B>"

    end if

    celdas = celdas & ofv.GenCelda("","VT ","","45%","","","",sInput,"si")

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

   sInput = "" 

    sInput = ofv.combosino(sForm,rs(10),rs(10).name)
    SINPUT = REPLACE(SINPUT,".all","." & sForm) 
    IF RS(10) = 1 THEN
       SINPUT = SINPUT & SIMAGS
    ELSE
       SINPUT = SINPUT & SIMAGN
    END IF 
    celdas = celdas & ofv.GenCelda("","VT","CENTER","10%","","","",sInput,"si")

   sInput = "" 


    sInput = ofv.combosino(sForm,rs(11),rs(11).name)

  '''  sInput = ofv.comboSiNoEx(sForm,rs(11),rs(11).name, "1", "0")
    SINPUT = REPLACE(SINPUT,".all","." & sForm) 

    IF RS(11) = 1 THEN
       SINPUT = SINPUT & SIMAGS
    ELSE
       SINPUT = SINPUT & SIMAGN
    END IF 
    celdas = celdas & ofv.GenCelda("","VT","CENTER","10%","","","",sInput,"si")

   sInput = "" 


    sInput = ofv.combosino(sForm,rs(12),rs(12).name)
    SINPUT = REPLACE(SINPUT,".all","." & sForm) 
    IF RS(12) = 1 THEN
       SINPUT = SINPUT & SIMAGS
    ELSE
       SINPUT = SINPUT & SIMAGN
    END IF 
    celdas = celdas & ofv.GenCelda("","VT","CENTER","10%","","","",sInput,"si")

   sInput = "" 


    sInput = ofv.combosino(sForm,rs(13),rs(13).name)
    SINPUT = REPLACE(SINPUT,".all","." & sForm) 
    IF RS(13) = 1 THEN
       SINPUT = SINPUT & SIMAGS
    ELSE
       SINPUT = SINPUT & SIMAGN
    END IF 
    celdas = celdas & ofv.GenCelda("","VT","CENTER","10%","","","",sInput,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=PROYRESP"
      formulario0 = ofv.GenForm(SFORM,"",sAccion,"","post","",celdas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",celdas,"si")
    end if


    filas = filas & ofv.GenRow("","","","","10","","",formulario0,"si")


    RS.MOVENEXT
    LOOP





    Celda = ofv.GenCelda("","abc colspan=7 style='border-bottom-style:RIDGE;border-bottom-width:2;border-bottom-color:#F7AA22;FONT-SIZE:5pt' ","center","100%","","","","&nbsp;","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")



    FilaS1 = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")





 if sopcion = "AT" THEN
     celdas = ofv.GenCelda("","","","1%","","","","&nbsp;","si")
 ELSE
     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYRESPBUS.asp?VASP=ES"
     sinput = sinput  & " ><b>&nbsp;AGREGAR&nbsp;</b></a>" 
     celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")
END IF

     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYSKR.asp?OPCION=RF"
     sinput = sinput  & " ><b>REFRESCAR</b></a>" 
     celdas = celdas & ofv.GenCelda("","tt","","10%","","","",sInput,"si")

     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYresp.asp"
     sinput = sinput  & " ><b>VOLVER</b></a>" 
     celdas = celdas & ofv.GenCelda("","tt","","10%","","","",sInput,"si")

  ''   celdas = celdas & ofv.GenCelda("","","","20%","","","","&nbsp;","si")
     celdas = celdas & ofv.GenCelda("","","","10%","","","","&nbsp;","si")

     sinput = sImagenO & "&nbsp;Originante" 
     celdas = celdas & ofv.GenCelda("","vt","","20%","","","",sInput,"si")
     sinput = sImagenI & "&nbsp;Interviniente" 
     celdas = celdas & ofv.GenCelda("","vt","","20%","","","",sInput,"si")
     sinput = SIMAGENA & "&nbsp;Afectado" 
     celdas = celdas & ofv.GenCelda("","vt","","20%","","","",sInput,"si")

     fila = ofv.GenRow("","","","","","","",celdas,"si")

     FilaS2 = ofv.GenTabla("","","","","","0","","4","0","","",fila,"si")



    Celda = ofv.GenCelda("","abc colspan=6","center","100%","","","","&nbsp;","si")
    FilaS =  ofv.GenRow("","","","","","","",Celda,"si")


  StrSql = "Select * From PROYRESP where codigo = '" & session("NCliente")  & "' " & XSEL2 & " ORDER BY TIPO,ORDEN "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","TTT colspan=6","","","","","",sImagen & "&nbsp;PERMISOS DE ACCESO","si")
  filas  = filaS & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","TTR colspan=5","center","","","","","GLOBAL/GERENCIAS/SECTORES","si")
  celdas = celdas & ofv.GenCelda("","TTR","center","","","","","ACCESO","si")

  filas = filas & ofv.GenRow("","","","","22","","",celdas,"si")



  DO UNTIL RS.EOF




    if RS(5) = "GL" then
       sInput = "<B><I>GLOBAL</I></B>"
       sInput1 = sImageG
    else

    if RS(5) = "SA" then
       sInput1 = sImageS

    ELSE
       sInput1 = sImageR
    END IF 

      strsql = "Select * From ECO where ECO = " & rs(3)
      set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
      if not rsniv.eof then sniv = rsniv(1)
      call ofv.cerrarconsulta(rsniv)
      sInput = "<B>" & sniv & "</B>"

    end if

    celdas = ofv.GenCelda("","VT ","center","1%","","","",SINPUT1,"si")
    celdas = celdas & ofv.GenCelda("","VT COLSPAN=4 ","","70%","","","",sInput,"si")

    SFORM = "Usuarios2" & RS(0)

   sInput = "" 

    sInput = ofv.combosino(sForm,rs(10),rs(10).name)
    SINPUT = REPLACE(SINPUT,".all","." & sForm) 
    IF RS(10) = 1 THEN
       SINPUT = SINPUT & SIMAGS
    ELSE
       SINPUT = SINPUT & SIMAGN
    END IF 
    celdas = celdas & ofv.GenCelda("","VT","CENTER","10%","","","",sInput,"si")



    if bModificar then
      sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=PROYRESP"
      formulario0 = ofv.GenForm(SFORM,"",sAccion,"","post","",celdas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",celdas,"si")
    end if


    filas = filas & ofv.GenRow("","","","","10","","",formulario0,"si")


    RS.MOVENEXT
    LOOP





    Celda = ofv.GenCelda("","abc colspan=6 style='border-bottom-style:RIDGE;border-bottom-width:2;border-bottom-color:#F7AA22;FONT-SIZE:5pt' ","center","100%","","","","&nbsp;","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")


    FilaS3 = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")







 if sopcion = "AT" THEN
     celdas = ofv.GenCelda("","","","1%","","","","&nbsp;","si")
 ELSE
     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYRESPBUS.asp?VASP=SC"
     sinput = sinput  & " ><b>&nbsp;AGREGAR&nbsp;</b></a>" 
     celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")
END IF

     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYresp.asp"
     sinput = sinput  & " ><b>VOLVER</b></a>" 
     celdas = celdas & ofv.GenCelda("","tt","","10%","","","",sInput,"si")

     celdas = celdas & ofv.GenCelda("","","","20%","","","","&nbsp;","si")

     sinput = sImageG & "&nbsp;Global" 
     celdas = celdas & ofv.GenCelda("","vt","","20%","","","",sInput,"si")
     sinput = sImageR & "&nbsp;Gerencia/Area" 
     celdas = celdas & ofv.GenCelda("","vt","","20%","","","",sInput,"si")
     sinput = SIMAGES & "&nbsp;Sector/Dpto" 
     celdas = celdas & ofv.GenCelda("","vt","","20%","","","",sInput,"si")

     fila = ofv.GenRow("","","","","","","",celdas,"si")

     FilaS4 = ofv.GenTabla("","","","","","0","","4","0","","",fila,"si")




    Celda = ofv.GenCelda("","","center","","","","",FilaS1,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")


    Celda = ofv.GenCelda("","","center","","","","",FilaS2,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","","center","","","","",FilaS3,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","","center","","","","",FilaS4,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"




  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""  


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","",Celda)

 




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML


FUNCTION GENREFRESH()

        XXXX = GENSECTORACC("999999","GL")

  StrSql = "Select * From PROYRESP where codigo = '" & session("NCliente")  & "' " & XSEL1 & " ORDER BY ORDEN "
  set rsXP = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF NOT RSXP.EOF THEN
     DO WHILE NOT RSXP.EOF 

        XXXX = GENSECTORACC(rsXP(3),"SA")
        XXXX = GENSECTORACC(rsXP(4),"RA")


        SFUNC = RSXP(15) & ""

        if SFUNC = "" then

           if RSXP(5) = "OR" then
              SFUNC = "3"
           ELSEif RSXP(5) = "AF" then
              SFUNC = "1"
           ELSE
              SFUNC = "2"
           END IF
           StrSql = "UPDATE PROYRESP SET FUNCION = '" & SFUNC & "' where ID = " & RSXP(0)
           CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

        end if



     RSXP.MOVENEXT
     LOOP

  END IF
  call ofv.cerrarconsulta(rsXP)

END FUNCTION





%>