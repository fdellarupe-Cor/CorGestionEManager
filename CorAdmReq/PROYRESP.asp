<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


 FUNCION = " onload=" & chr(34)
 funcion = funcion & " if (window.name=='Principal') {"
 funcion = funcion & " document.all.UNO.style.display='none';} " & chr(34)



  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

  shtml = replace(shtml,"<body ","<body " & funcion)



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

 if sopcion = "AU" THEN
    RESP = request.querystring("RESP")
    USR = request.querystring("NUSUARIO")
    StrSql = "UPDATE PROYRESP SET USRAFC = '" & USR & "' WHERE ID = " & RESP
    CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
 END IF

 XSEL = ""

 if sopcion = "AT" THEN
    XSEL = "' AND USRAFC <> 'NO"
 END IF

    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenA = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
    sImagenC = "<B><FONT FACE=WEBDINGS COLOR=#888888 SIZE=3>&#0128;</FONT></B>"
    sImageR = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>=</FONT></B>"
    sImageS = "<B><FONT FACE=WEBDINGS COLOR=#6699CC SIZE=3>4</FONT></B>"
    sImageG = "<B><FONT FACE=WEBDINGS COLOR=#007700 SIZE=3>&#0252;</FONT></B>"
    sImagenM = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=2>&#0243;</FONT></B>"
    sImageF = "<B><FONT FACE=WEBDINGS COLOR=#005500 SIZE=3>&#0151;</FONT></B>"


  StrSql = "Select * From PROYRESP where codigo = '" & session("NCliente") & XSEL & "' AND SKR <> 'A' ORDER BY ORDEN "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

     GENSECTOROR()

     StrSql = "Select * From PROYRESP where codigo = '" & session("NCliente") & XSEL & "' AND SKR <> 'A'  ORDER BY ORDEN "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  END IF




    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","ttt colspan=5","","","","","",sImagen & "&nbsp;RESPONSABLES - GERENCIAS/SECTORES/USUARIOS AFECTADOS","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","","center","","","","","&nbsp;","si")
  celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
  celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
  celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
  celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  POSI = 0

  DO UNTIL RS.EOF

  POSI = POSI + 1

    'Eliminar
    if bSuprimir AND RS(5) = "AF" then

       IF RS(7) <> "NO" THEN
          StrSql = "Select ID From PROYTAREAS WHERE USRRESP = '" & RS(7) & "' AND CODIGO = '" & session("NCliente") & "' "
          set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          IF RS2.EOF THEN
             SELIM = "S" 
          ELSE
             SELIM = "N" 
          END IF
       ELSE
          SELIM = "S" 
       END IF


    else
       SELIM = "N" 
    end if


       IF SELIM = "S" THEN  
         sHref = "asp/EliminarRegistro.asp?ID=" & rs(0)
         sHref = sHref & "&Tabla=PROYRESP&volver=" & vuelta & "&Eliminar=NSNC"
         CeldaS = ofv.BotonEliminar(" RESPONSABLE ",sHref)
         CeldaS = REPLACE(CeldaS,"<TD","<TD WIDTH=1% ROWSPAN=2 ")
       ELSE
         CeldaS = ofv.GenCelda("","UT ROWSPAN=2 ","","1%","","","","&nbsp;","si")
       END IF





    SELECT CASE RS(5)
           CASE "OR"
                 sInput = sImagenO
           CASE "AN"
                 sInput = sImagenI
           CASE "AF"
                 sInput = sImagenA
           CASE "UC"
                 sInput = sImagenC
    END SELECT

     celdas = CELDAS & ofv.GenCelda("","vt ROWSPAN=2 BGCOLOR=#FFFFFF style='border-RIGHT-style:SOLID;border-RIGHT-width:1;border-RIGHT-color:#6699CC;' ","","1%","","","",sInput,"si")



    if RS(7) = "NO" then
       sInput = "<B><I>USUARIO: SIN AFECTAR</I></B>"

       sinput2 = "<a CLASS=CC style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "none" & chr(34) & ";' class=tt href=PROYRESPAFT.asp?RESP=" & RS(0) & "&SECT=" & RS(3) 
       sinput2 = sinput2  & " TITLE='AFECTAR' ><b>" & sInput & "</b></a>" 
       SCOLOR =  "ata bgcolor=#800000 "
    else

      strsql = "Select * From USUARIOS where ID = '" & rs(7) & "' "
      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsniv.eof then sniv = rsniv(1)
      call ofv.cerrarconsulta(rsniv)
      sInput = "<B>" & sniv & "</B>"

 IF RS(5) = "UC" THEN

       sinput2 = sInput

 ELSE

 if sopcion = "AT" THEN
       sinput2 = "<a  CLASS=CC style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "none" & chr(34) & ";' class=tt href=PROYTASKPROP.asp?OPCION=AT&USR=" & RS(7) 
       sinput2 = sinput2  & " TITLE='ASIGNAR'  ><b>" & sInput & "</b></a>" 
       SCOLOR =  "ata bgcolor=#f7aa22 "
 ELSE
       sinput2 = "<a  CLASS=CC style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "none" & chr(34) & ";' class=tt href=PROYRESPAFT.asp?RESP=" & RS(0) & "&SECT=" & RS(3) 
       sinput2 = sinput2  & " TITLE='MODIFICAR'  ><b>" & sInput & "</b></a>" 
       SCOLOR =  "ata bgcolor=#008000 "

 END IF 

 END IF


    end if


    celdas = celdas & ofv.GenCelda("","vt ","","50%","","","",sInput2,"si")

    celdas = celdas & ofv.GenCelda("","vT  ","","1%","","","",sImageR,"si")

    if RS(4) = 0 then
      celdas = celdas & ofv.GenCelda("","vt","","40%","","","","<FONT COLOR=#336699 >SIN AFECTAR</FONT>","si")

    else

     StrSql = "Select * From ECO where ECO = " & RS(4)
     set rs2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rs2.eof then
        sInput = rs2(1)
     ELSE
        sInput = "NO DETERMINADO"
     END IF
     celdas = celdas & ofv.GenCelda("","vt","","40%","","","","<FONT COLOR=#336699 >" & sInput & "</FONT>","si")


    end if

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




        sinputK = ""
        sinputt = ""

        SFORM = "F" & POSI

    if bModificar AND RS(5) <> "UC" then
      strsql = "Select * From PROYFUNCIONES Order By DSCFULL"
      sInput = ofv.GenerarCombo(SFORM,rs(15).name,rs(15),strsql,cn,0,2,"")
      sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=PROYRESP"
      sInput = ofv.GenForm(SFORM,"",sAccion,"","post","",sInput,"si")
      sinputt = "<div id=AA" & POSI & " name=xxx style='display:none;' >" &  sInput & "</div>"

        sinputK = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  href='javascript:void(0);' onclick='AA" & POSI & ".style.display=" & chr(34) & "block" & chr(34) & ";AB" & POSI & ".style.display=" & chr(34) & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  TITLE='Modificar funcion'  ><b>" & sImagenM & "</b></a>" 

    end if

      sInput = "S/D"
      strsql = "Select * From PROYFUNCIONES where ID = " & rs(15)
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then sInput = rsniv(2)
        call ofv.cerrarconsulta(rsniv)

        sInput = "<div id=AB" & POSI & " name=aaa style='display:block;' ><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR=#336699 SIZE=1><U>" &  sInput & "</U></FONT>&nbsp;" &  sInputK & "</span></div>" &  sInputt

     celdas = ofv.GenCelda("","UT","","40%","","","",sInput,"si")



    celdas = CELDAS & ofv.GenCelda("","vT  ","","1%","","","",sImageS,"si")

    if RS(3) = 0 then
      celdas = CELDAS & ofv.GenCelda("","vt","","40%","","","","<FONT COLOR=#336699 >SIN AFECTAR</FONT>","si")

    else

     StrSql = "Select * From ECO where ECO = " & RS(3)
     set rs2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rs2.eof then
        sInput = rs2(1)
     ELSE
        sInput = "NO DETERMINADO"
     END IF
     celdas = CELDAS & ofv.GenCelda("","vt","","40%","","","","<FONT COLOR=#336699 >" & sInput & "</FONT>","si")


    end if

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    celdas = ofv.GenCelda("","VT COLSPAN=5 style='border-bottom-style:SOLID;border-bottom-width:1;border-bottom-color:#6699CC;FONT-SIZE:5pt' ","center","","","","","&nbsp;","si")

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    RS.MOVENEXT
    LOOP





    Celda = ofv.GenCelda("","abc colspan=5","center","100%","","","","&nbsp;","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")

    FilaS = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")

    Celda = ofv.GenCelda("","","center","","","","",FilaS,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")


 if sopcion = "AT" THEN
     celdas = ofv.GenCelda("","","","1%","","","","&nbsp;","si")
     celdas = celdas & ofv.GenCelda("","","","1%","","","","&nbsp;","si")
 ELSE
     celdas = ""

     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYRESPBUS.asp?VASP=ES"
     sinput = sinput  & " ><b>&nbsp;afectar&nbsp;sector&nbsp;</b></a>" 
 '    celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")

          sInput  = ofv.GenerarInput("boton","AFECTAR","submit","","10","bt","")
          sAccion = "PROYRESPBUS.asp?VASP=ES"
          btcel   = ofv.GenCelda("","UT TITLE='AFECTAR SECTOR.' ","","","","","",sInput,"si")
          celdas  = celdas & ofv.GenForm("init","",sAccion,"","post","",btcel,"si")

     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYSKR.asp"
     sinput = sinput  & " ><b>&nbsp;PERMISOS&nbsp;</b></a>" 
   '  celdas = celdas & ofv.GenCelda("","tt","","10%","","","",sInput,"si")

          sInput  = ofv.GenerarInput("boton","PERMISOS","submit","","10","bt","")
          sAccion = "PROYSKR.asp"
          btcel   = ofv.GenCelda("","UT TITLE='PERMISOS DEL PROYECTO.' ","","","","","",sInput,"si")
          celdas  = celdas & ofv.GenForm("init","",sAccion,"","post","",btcel,"si")

END IF

     sinput = "&nbsp;"

     celdas = celdas & ofv.GenCelda("UNO","ut","","10%","","","",sInput,"si")
     celdas = celdas & ofv.GenCelda("","","","80%","","","","&nbsp;","si")

     fila = ofv.GenRow("","","","","","","",celdas,"si")

     Fila = ofv.GenTabla("","","","","","0","","4","0","","",fila,"si")

    Celda = ofv.GenCelda("","","center","","","","",Fila,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"



  'Botonera
  Agregar = ""
  Buscar  = ""

    sAccion = "PROYPROP.asp"
    Celda   = ofv.BotonVolver(sAccion)     


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","",Celda)



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML



%>