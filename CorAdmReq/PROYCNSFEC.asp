<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0






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


 XMOD = request.querystring("XMOD")
 XDST = request.querystring("XDST")

 XSET = request.querystring("XSET")
 XOPT = request.querystring("XOPT")
 sRESP = request.querystring("RESP")

 sopcion = request.querystring("opcion")

 sACCION = request.querystring("ACCION")


vNuevoValor = clng(DATE()) - nFecini



  sHTML = ofv.FormHeader(replace(session("FormN"),"_",FUNCION))



    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=3>R</FONT></B>"
 

    sImagenFF = "<B><FONT FACE=WEBDINGS COLOR=#008000 SIZE=4>4&#0166;</FONT></B>"
    sImagenF1 = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=4>4</FONT></B>"

IF SOPCION = "MOT" OR SOPCION = "MOR" THEN
   


   IF SOPCION = "MOT" THEN
      SESSION("FECOPC") = "1"
      SESSION("FECDES") = 0
      SESSION("FECHAS") = 99999 
   ELSE
      SESSION("FECOPC") = "2"
      SESSION("FECDES") = CLNG(DATE()) -  clng(cdate("01-01-1900"))
      SESSION("FECHAS") = CLNG(DATE()) -  clng(cdate("01-01-1900"))
   END IF

   SOPCION = ""
   sACCION = ""

END IF

IF SOPCION = "SIN" OR SOPCION = "SNI"  OR SOPCION = "SSI"  THEN
   


   IF SOPCION = "SIN" THEN
      SESSION("SELOPC") = "0"
   ELSEIF SOPCION = "SSI" THEN
      SESSION("SELOPC") = "2"
   ELSE
      SESSION("SELOPC") = "1"
   END IF

   SOPCION = ""
   sACCION = ""

END IF

IF SACCION = "AF" THEN
   


   IF SOPCION = "MFD" THEN

      SESSION("FECDES") = REQUEST.FORM("FECHA")

   ELSE

      SESSION("FECHAS") = REQUEST.FORM("FECHA")

   END IF

   SOPCION = ""
   sACCION = ""

END IF


    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","cc colspan=3","","","","","",sImagen & "&nbsp;<BIG><B>MODIFICACION RANGO DE SELECCION</B></BIG>","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")


  celdas = ofv.GenCelda("","ABC colspan=3","","","","","","&nbsp;","si")
  filas  = FILAS & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","vt","","","","","","SELECCION:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")


     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href='javascript:void(0);' "
     sinput = sinput  & "  onclick=" & chr(34) & "document.all.S1.style.display='block';"
     sinput = sinput  & "document.all.S2.style.display='block';"
     sinput = sinput  & "document.all.S3.style.display='block';" & chr(34) 
     sinput = sinput  & "   TITLE='MODIFICAR SELECCION DE INCIDENCIAS'  ><b>&nbsp;" & sImagenFF & "</b></a>" 

    if SESSION("SELOPC") = "1" then
  '    sinput = sinput  & "&OPCION=SIN"
      sInput2 = "NO INCLUYE INCIDENCIAS"
    elseif SESSION("SELOPC") = "2" then
   '   sinput = sinput  & "&OPCION=SSI"
      sInput2 = "SOLO INCIDENCIAS"
    else
   '   sinput = sinput  & "&OPCION=SNI"
      sInput2 = "INCLUYE INCIDENCIAS"
    end if

  '   sinput = sinput  & "   TITLE='MODIFICAR SELECCION DE INCIDENCIAS'  ><b>&nbsp;" & sImagenFF & "</b></a>" 

    celdas = celdas & ofv.GenCelda("","vt","","60%","","","",sInput2 & sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYCNSFEC.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP
     sinput = sinput  & "&OPCION=SNI"
     sinput = sinput  & "   TITLE='MODIFICAR SELECCION DE INCIDENCIAS'  ><b>&nbsp;" & sImagenF1 & "NO INCLUYE INCIDENCIAS</b></a>" 


    celdas = ofv.GenCelda("S1","vt COLSPAN=3 style='display:none;' ","RIGHT","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYCNSFEC.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP
     sinput = sinput  & "&OPCION=SIN"
     sinput = sinput  & "   TITLE='MODIFICAR SELECCION DE INCIDENCIAS'  ><b>&nbsp;" & sImagenF1 & "INCLUYE INCIDENCIAS&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></a>" 

    celdas = ofv.GenCelda("S3","vt COLSPAN=3 style='display:none;' ","RIGHT","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYCNSFEC.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP
     sinput = sinput  & "&OPCION=SSI"
     sinput = sinput  & "   TITLE='MODIFICAR SELECCION DE INCIDENCIAS'  ><b>&nbsp;" & sImagenF1 & "SOLO INCIDENCIAS&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></a>" 

    celdas = ofv.GenCelda("S2","vt COLSPAN=3 style='display:none;' ","RIGHT","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




    celdas = ofv.GenCelda("","vt","","","","","","OPCION:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")


     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYCNSFEC.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP



    if SESSION("FECOPC") = "1" then
      sinput = sinput  & "&OPCION=MOR"
      sInput2 = "TODOS"
    else
      sinput = sinput  & "&OPCION=MOT"
      sInput2 = "POR RANGO"
    end if

     sinput = sinput  & "   TITLE='MODIFICAR RANGO DE SELECCION'  ><b>&nbsp;" & sImagenFF & "</b></a>" 

    celdas = celdas & ofv.GenCelda("","vt","","60%","","","",sInput2 & sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    if SESSION("FECOPC") = "2" then

    

    celdas = ofv.GenCelda("","vt","","","","","","FECHA DESDE:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYCNSFEC.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP  & "&OPCION=MFD"
     sinput = sinput  & "   TITLE='MODIFICAR FECHA DESDE'  ><b>&nbsp;" & sImagenI & "</b></a>" 

       dFecha = Fechador(SESSION("FECDES"))




      sInput = dFecha & SINPUT


    celdas = celdas & ofv.GenCelda("","vt","","60%","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")






    celdas = ofv.GenCelda("","vt","","","","","","FECHA HASTA:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYCNSFEC.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP  & "&OPCION=MFH"
     sinput = sinput  & "   TITLE='MODIFICAR FECHA HASTA'  ><b>&nbsp;" & sImagenI & "</b></a>" 

       dFecha = Fechador(SESSION("FECHAS"))

      sInput = dFecha & SINPUT


    celdas = celdas & ofv.GenCelda("","vt","","60%","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  END IF



    FilaS = ofv.GenTabla("","AAA style='border-BOTTOM-style:ridge;border-BOTTOM-width:2;border-BOTTOM-color:#f7aa22;' ","","100%","","0","","0","0","","",filas,"si") & "<BR>"

'''''''''''''''''''''
 xSTS = 0

IF SOPCION = "MFD" OR SOPCION = "MFH" THEN xSTS = 1

XFSR = REQUEST.FORM("FECHA")

IF XFSR = "" THEN XFSR = CLNG(DATE()) -  clng(cdate("01-01-1900"))

    IF xSTS = 1 THEN

  cniv = request.querystring("Cniv")
  cnom = request.querystring("Cnom")
  cx = request.querystring("Cx")
  cA = request.querystring("CA")

  if cniv = "" then 
     cniv = session("ENC7")
     cnom = session("ENC8")
     cx = session("ENC9")
     cA = session("ENCA")
  else
     session("ENC7") = cniv
     session("ENC8") = cnom
     cx = cdbl(cx)
     session("ENC9") = cx
     session("ENCA") = cA
  end if






  posi = 0
    sForm = "form" & posi
       sclase = "CC" 

     Celda = ofv.GenCelda("","VT","","15%","","","","FECHA:","si")



    if clng(XFSR) = 0 then
       sInput1 = "------------"
    else
       sInput1 = cdate(clng(XFSR) + clng(cdate("01-01-1900")) )
       swd = datepart("w",sInput1)
       sInput1 = Fechador(XFSR)
    END IF


     Celda = Celda & ofv.gencelda("","VT colspan=3 ","","80%","","","",sInput1,"si")



    Fila = ofv.genrow("","","","","","","",Celda,"si")

  ''  filas = filas & ofv.GenTabla("","","","50%","","0","","0","0","","",Fila,"si")

    'Descripcion
    CELDA = ""

    sanio = datepart("yyyy",date())

    sanioC = datepart("yyyy",date())
    smes = datepart("m",date())

    SMESC = SMES

    if not isnull(XFSR) then

       if clng(XFSR) > 0 then

          smesc = cdate(clng(XFSR) + clng(cdate("01-01-1900")) )
          smesc = datepart("m",smesc)

       end if 


    end if

    IF CA <> "" THEN 
       smesc = request.form("smesc")
       sanioC = request.form("sanioc")

       IF (LEN(SANIOC&"")=0) OR CSTR(SANIOC) = "0" THEN SANIOC = SANIO

 '      response.write "smesc " & smesc & " ca " & ca
       if isnull(smesc) then  SMESC = ca
       if len(smesc) = 0 then  SMESC = ca
       ca =  SMESC
      
    else
       ca = smesc

    end if  

    if smesc = 0 then smesc = smes

  '     response.write "smesc " & smesc & " ca " & ca

       session("ENCA") = smesc

'    if cint(smesc) = cint(smes) then
'       sanioc = sanio
'    elseif cint(smesc) < cint(smes) then
'       sanioc = sanio + 1
'    else
'       sanioc = sanio
'    end if

    
      Celda = ofv.GenCelda("","VT","","15%","","","","MES:","si")

    'codigo

     Strsql = "Select * From meses Order By mes"
     sInput = ofv.GenerarCombo(sForm & "A","smesc",smesc,Strsql,cn2,1,2,"sb0")
     Celda = Celda & ofv.gencelda("","VT","","45%","","","",sInput,"si")

     Celda = Celda & ofv.GenCelda("","VT","","15%","","","","AÑO:","si")


 ''    RESPONSE.WRITE "A " & SANIOC & " X " & SANIO & "<BR>"



        sInput = "<font class=at><select name=SANIOC onchange=submit()>"

        FOR XYY=-2 TO 2


            sInput = sInput & "<option value=" & (CINT(SANIO) + XYY)
            if CINT(CINT(SANIO) + XYY) = CINT(SANIOC) then sInput = sInput & " selected"
            sInput = sInput & " >" & (CINT(SANIO) + XYY)

        NEXT

        sInput = sInput & "</select></font>"




     
'''     sinput = sanioc 
     Celda = Celda & ofv.gencelda("","VT","","20%","","","",sInput,"si")

    Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")


    sAccion = "PROYCNSFEC.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST & "&XOPT=" & XOPT & "&RESP=" & sresp & "&OPCION=" & SOPCION & "&ACCION=AM" & "&cniv=" & cniv & "&cnom=" &  cnom & "&cx=" & cx & "&ca=" & smesc
    Formulario = ofv.genform(sForm & "A","",sAccion,"","post","",Fila,"si")

    filas = filas & ofv.GenTabla("","AAA style='border-style:ridge;border-width:2;border-color:#f7aa22;'","","50%","","0","","0","0","","",Formulario,"si")


    dim afech(33,1)
    dim afechs(42,1)

    CELDA = ""

    sfecds = "01-" & cstr(smesc) & "-" & cstr(sanioc)

    'sfecd = clng(cdate(sfecds))
    sfecd = clng(dateserial(sanioc,smesc,1))

    'swd = datepart("w",date(sfecds))
     swd = datepart("w",dateserial(sanioc,smesc,1))
     
    sfechs = "28-" & cstr(smesc) & "-" & cstr(sanioc)

    'sfech = clng(date(sfechs))
     sfech = clng(dateserial(sanioc,smesc,28))
 '   response.write " fd " & sfecd & " fh " & sfech

    yy = 0

    for ii = 1 to 3
        sfecha = sfech + ii
        sfecha = cdate(clng(sfecha))
        smess = datepart("m",sfecha)
        if cint(smesc) <> cint(smess) then
           exit for
        else
           yy = ii
        end if

    next  

    sfech = sfech + yy

 '  response.write " wd " & swd & " fd " & sfecd & " fh " & sfech
    srest = clng(cdate("01-01-1900"))
   ' response.write  sfech &" -"&sfecd 
    topefech = sfech - sfecd 

    yy = 0



    wswd = swd

    for ii = 0 to topefech
        
        afech(ii+1,0) = wswd
        wswd = wswd + 1
        if wswd > 7 then wswd = 1        

        afech(ii+1,1) = sfecd + ii


    next


    for ii = 1 to (cint(afech(1,0)) - 1)
        afechs(ii,0) = "NO"
        afechs(ii,1) = 0
    next

    yy = 1

 '   RESPONSE.WRITE "II " & II & " TOPE " & (ii + topefecH ) & " TOPEFECH " & topefecH
    for ii = ii to (ii + topefecH )
        afechs(ii,0) = yy
        afechs(ii,1) = afech(yy,1)
        yy = yy + 1
    next

'    RESPONSE.WRITE "II " & II

    for ii = ii to 42
        afechs(ii,0) = "NO"
        afechs(ii,1) = 0
    next

    CELDA = ""
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;'","CENTER","10%","","","","D","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;'","CENTER","10%","","","","L","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;'","CENTER","10%","","","","M","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;'","CENTER","10%","","","","M","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;'","CENTER","10%","","","","J","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;'","CENTER","10%","","","","V","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;'","CENTER","10%","","","","S","si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")

    CELDA = ""
    yy = 0

    for ii = 1 to 43

        if yy = 7 then
           yy = 1
           Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
           CELDA = ""
        else
           yy = yy + 1 
        end if

        if ii > 42 then exit for  

        if yy = 1 or yy = 7 then
           sclase = "tt3 style='border-style:outset;border-width:2;'" 
        else
           sclase = "TT style='border-style:outset;border-width:2;'" 
        end if 

        if afechs(ii,0) = "NO" then
           sinput = "&nbsp;"
           sclase = "CC " 
           Celda = Celda & ofv.GenCelda("","VT","CENTER","10%","","","",sinput,"si")
        else
           sinput1 = afechs(ii,0)

           sinput1 = cstr(sinput1)

           if len(sinput1) = 1 then sinput1 = "0" & sinput1 

           sinput1 = " " & sinput1 & " " 

           sAccion = "PROYCNSFEC.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST & "&XOPT=" & XOPT & "&RESP=" & sresp & "&OPCION=" & SOPCION & "&ACCION=AF"

           t1 = ofv.GenerarInput("bvo1",sInput1,"submit","","10",sclase,"")
           t1 = t1 & ofv.GenerarInput("FECHA",(afechs(ii,1) - srest),"hidden","","2","dt","")
           t1 = ofv.GenCelda("","VT","center","10%","","","",t1,"si")
           t1 = ofv.GenForm("feac" & ii,"",sAccion,"","post","",t1,"si") 
           Celda = Celda & t1

        end if


    next

'''


   filas = filas & ofv.GenTabla("","AAA style='border-style:ridge;border-width:2;border-color:#f7aa22;'","","50%","","0","","0","0","","",FILA,"si")

 
 END IF


'''''''''''''''''''''


    celdas = ""

       sinput = "<a class=CC  href=PROYCNSLST.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST & "&XOPT=" & XOPT & "&RESP=" & sresp
       sinput = sinput  & " ><b><FONT FACE=WEBDINGS size=2 >:</FONT>  VOLVER</b></a>" 
       celdas = celdas & ofv.GenCelda("","VT","RIGHT","","10%","","",sInput,"si")

    cel = ofv.GenRow("","","","","","","",celdas,"si")
    FilaS =  filaS & "<BR><BR>" & ofv.GenTabla("","aaa style='border-top-style:ridge;border-top-width:2;border-top-color:#f7aa22;' ","","100%","","0","","4","0","","",cel,"si")


    Celda = ofv.GenCelda("","","center","","","","",FilaS,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")


    sHTML = sHTML & ofv.GenTabla("","aa ","","80%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"

 
  IF session("ModMain2") = "CL" THEN
     sinput = "<a  class=tt href=Clientes.asp"
     sinput = sinput  & " ><b>&nbsp;Volver&nbsp;a&nbsp;Cliente&nbsp;</b></a>" 
     celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")

  ELSE
     sinput = "&nbsp;"
     celdas =  ofv.GenCelda("","","","1","","","",sInput,"si")
  END IF


  celdas = celdas & ofv.GenCelda("","","","","","","","&nbsp;","si")
  filas = ofv.GenRow("","","","","","","",celdas,"si")
  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si")

'  end if

' END IF



 ' call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML




Function Fechador(wkfecha)


    
      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))
    
      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function

%>