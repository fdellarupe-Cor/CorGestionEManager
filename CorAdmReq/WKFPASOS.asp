<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(session("FormN"))


  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn5)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

 sopcion = request.querystring("opcion")
 WKF = request.querystring("WKF")

 if WKF = "" then
    WKF = session("WKF")
 else
    session("WKF") = WKF
 end if 

 if sopcion = "SI" then
    GENWKFPASO()

 end if
    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>&#0143;</FONT></B>"
    sImagenI = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=2>&#0243;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WebDINGS COLOR=#008000 SIZE=4>&#0194;</FONT></B>"
    sImagenPI = "<B><FONT FACE=WebDINGS COLOR=#99bbcc SIZE=3>4</FONT></B>"
    sImagenPF = "<B><FONT FACE=WebDINGS COLOR=#0000aa SIZE=3><</FONT></B>"
    sImagenPC = "<B><FONT FACE=WebDINGS COLOR=#6699cc SIZE=3>.</FONT></B>"
    sImagenCP = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>v</FONT></B>"
    sImagenI1 = "<B><FONT FACE=WEBDINGS COLOR=#000000 SIZE=3>&#0128;</FONT></B>"
    sImagenI2 = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=4>N</FONT></B>"
    sImagenI3 = "<B><FONT FACE=WINGDINGS COLOR=#cc0000 SIZE=3>&#0232;</FONT></B>"



  StrSql = "Select * From WKFCIRCUITO WHERE WKFCOD = '" & WKF & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
'     IF RS(3) THEN
'        bAgregar = FALSE
'        bModificar = FALSE
'        bSuprimir = FALSE
'     END IF

    IDWKF = RS(0)  

    SFORM = "Usuarios"

    sInput = "CIRCUITO DE WORKFLOW"
    celdas = ofv.GenCelda("","TTT","","20%","","","",sInput,"si")

'    sInput = RS(1) & " - <B>" & RS(2) & "</B>"
'    celdas = celdas & ofv.GenCelda("","cc","","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")
    sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si")

    filas = ""

    celdas = ofv.GenCelda("","UT","","20%","","","","&nbsp;CODIGO:","si")

    sInput = RS(1)

    celdas = celdas & ofv.GenCelda("","UT","","40%","","","",sInput,"si")


    celdas = celdas & ofv.GenCelda("","UT","","20%","","","","&nbsp;CLASE:","si")


        sinputK = ""
        sinputt = ""

    scomK = "E" & chr(9) & "M"
    scomD = "EFECTIVO" & chr(9) & "MODELO"

    if bModificar then
        sInput = ofv.gencomboFX(sForm,rs(5).name,rs(5),scomK,scomD," ")
    else
        sInput = ofv.GenerarInput("",rs(5),"readonly","","70","dt","")
    end if



'    if bModificar  then
'        sInput = "<font class=at><select name=" & rs(5).name & " onchange=submit()>"
'        sInput = sInput & "<option value=E "
'        if rs(5) = "E" then sInput = sInput & "selected"
'        sInput = sInput & " >EFECTIVO<option value=M "
'        if rs(5) = "M" then sInput = sInput & "selected"
'        sInput = sInput & " >MODELO</select></font>"
'        sinputt = "<div id=A21 name=xxx style='display:none;' >" &  sInput & "</div>"

'        sinputK = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
'        sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
'        sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
'        sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
'        sinputK = sinputK  & "none" & chr(34) & ";' "
'        sinputK = sinputK  & "  href='javascript:void(0);' onclick='A21.style.display=" & chr(34) & "block" & chr(34) & ";A22.style.display=" & chr(34) & "none" & chr(34) & ";' "
'        sinputK = sinputK  & "  TITLE='Modificar'  ><b>" & sImagenI & "</b></a>" 

'    end if


'        if rs(5) = "E" then sInput = "EFECTIVO"
'        if rs(5) = "M" then sInput = "MODELO"

'        sInput = "<div id=A22 name=aaa style='display:block;' ><span>" &  sInput & "&nbsp;" &  sInputK & "</span></div>" &  sInputt


    celdas = celdas & ofv.GenCelda("","UT","","40%","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","UT","","","","","","&nbsp;DESCRIPCION:","si")

   ''' sInput = RS(2)

       if bModificar then
          sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",SFORM,"70","dt","")
       else
          sInput = ofv.GenerarInput("",rs(2),"readonly","","70","dt","")
       end if

    celdas = celdas & ofv.GenCelda("","UT","","","","","",sInput,"si")


    celdas = celdas & ofv.GenCelda("","UT","","","","","","&nbsp;TIPO:","si")

        sinputK = ""
        sinputt = ""

    if bModificar then
      '''  sInput = ofv.gencomboFX(sForm,rs(5).name,rs(5),scomK,scomD," ")
        strsql = "Select * From WKFTIPOS Order By DESCRIPCION"
        sInput = ofv.GenerarCombo(SFORM,rs(6).name,rs(6),strsql,cn,1,2,"")

    else
        sInput = ofv.GenerarInput("",rs(6),"readonly","","70","dt","")
    end if


'    if bModificar then
'      strsql = "Select * From WKFTIPOS Order By DESCRIPCION"
'      sInput = ofv.GenerarCombo("Usuarios",rs(6).name,rs(6),strsql,cn,1,2,"")
'      sinputt = "<div id=A31 name=xxx style='display:none;' >" &  sInput & "</div>"

'        sinputK = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
'        sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
'        sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
'        sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
'        sinputK = sinputK  & "none" & chr(34) & ";' "
'        sinputK = sinputK  & "  href='javascript:void(0);' onclick='A31.style.display=" & chr(34) & "block" & chr(34) & ";A32.style.display=" & chr(34) & "none" & chr(34) & ";' "
'        sinputK = sinputK  & "  TITLE='Modificar'  ><b>" & sImagenI & "</b></a>" 

'    end if

'      sInput = "S/D"
'      strsql = "Select * From WKFTIPOS where TIPO = '" & rs(6) & "' "
'        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
'        if not rsniv.eof then sInput = rsniv(2)
'        call ofv.cerrarconsulta(rsniv)

'        sInput = "<div id=A32 name=aaa style='display:block;' ><span>" &  sInput & "&nbsp;" &  sInputK & "</span></div>" &  sInputt


    celdas = celdas & ofv.GenCelda("","UT","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





    celdas = ofv.GenCelda("","UT","","","","","","&nbsp;INICIO:","si")

        sinputK = ""
        sinputt = ""

    scomK = "A" & chr(9) & "S"
    scomD = "AUTOMATICO" & chr(9) & "A SOLICITUD"

    if bModificar then
        sInput = ofv.gencomboFX(sForm,rs(4).name,rs(4),scomK,scomD," ")
    else
        sInput = ofv.GenerarInput("",rs(4),"readonly","","70","dt","")
    end if


'    if bModificar  then
'        sInput = "<font class=at><select style='border-style:outset;border-width:5;' name=" & rs(4).name & " onchange=submit() >"
'        sInput = sInput & "<option value=A "
'        if rs(4) = "A" then sInput = sInput & "selected"
'        sInput = sInput & " >AUTOMATICO<option value=S "
'        if rs(4) = "S" then sInput = sInput & "selected"
'        sInput = sInput & " >A SOLICITUD</select></font>"
'        sinputt = "<div id=A11 name=xxx style='display:none;' >" &  sInput & "</div>"

'        sinputK = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
'        sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
'        sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
'        sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
'        sinputK = sinputK  & "none" & chr(34) & ";' "
'        sinputK = sinputK  & "  href='javascript:void(0);' onclick='A11.style.display=" & chr(34) & "block" & chr(34) & ";A12.style.display=" & chr(34) & "none" & chr(34) & ";' "
'        sinputK = sinputK  & "  TITLE='Modificar'  ><b>" & sImagenI & "</b></a>" 

'    end if


'        if rs(4) = "A" then sInput = "AUTOMATICO"
'        if rs(4) = "S" then sInput = "A SOLICITUD"

'        sInput = "<div id=A12 name=aaa style='display:block;' ><span>" &  sInput & "&nbsp;" &  sInputK & "</span></div>" &  sInputt



    celdas = celdas & ofv.GenCelda("","UT","","","","","",sInput,"si")

    celdas = celdas & ofv.GenCelda("","UT","","","","","","&nbsp;","si")

    celdas = celdas & ofv.GenCelda("","UT","","","","","","&nbsp;","si")

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(IDWKF," ","*") & "&volver=" & vuelta & "&Tabla=WKFCIRCUITO"
      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if

    filas = ofv.GenRow("","","","","","","",formulario0,"si")

  sHTML = sHTML  & ofv.GenTabla("","","","100% STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;' ","","0","","1","-1","","",filas,"si") & "<BR>"


  END IF



    celdas = ""



  'Encabezados
  formulario =  ofv.GenCelda("","ttR colspan=7","center","","","","","PASOS","si")
  formulario = ofv.GenRow("","","","","22","","",formulario,"si")


  StrSql = "Select * From WKFPASOS WHERE WKFCOD = '" & WKF & "' ORDER BY ORDEN"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF RS.EOF THEN
     GENWKFPASOIF()
     StrSql = "Select * From WKFPASOS WHERE WKFCOD = '" & WKF & "' ORDER BY ORDEN"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  END IF


  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    sform = "WKF" & rs(0)

    if rs(5) then
       sc = "CC"
    else
       sc = "TT"
    end if

       sc = "UT"

     sinput = "<a style='text-decoration:none;color:#ffffff;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";'  href=PROYtask.asp?opcion=SBT&TSK="
     sinput = sinput  & " TITLE='SUBIR TAREA' ><b>" & sImagen1 & "</b></a>" 


    celdas = ofv.GenCelda("","Ut  bgcolor=#6699cc ","","","","","",sinput,"si")



       sInput = rs(3) & ".&nbsp;"
       celdas = celdas & ofv.GenCelda("",sc,"right","","","","",sInput,"si")



       if bModificar then
          sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sform,"60","dt","")
       else
          sInput = ofv.GenerarInput("",rs(4),"readonly","","60","dt","")
       end if
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")


        if rs(6)  then     sInput = "" & sImagenPI & ""
        if rs(9)  then     sInput = "" & sImagenPF & ""
        if rs(7) OR RS(8) then     sInput = "" & sImagenPC & ""



    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")



      sInput = "" & sImagenCP & ""
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

    if  bModificar AND RS(3) > 10 AND RS(3) < 999 then


        scomK = "U" & chr(9) & "R"
        scomD = "NORMAL" & chr(9) & "RECURRENTE"

        if bModificar then
            sInput = ofv.gencomboFX(sForm,rs(12).name,rs(12),scomK,scomD," ")
        else
            sInput = ofv.GenerarInput("",rs(12),"readonly","","70","dt","")
        end if

    ELSE

        if rs(12) = "U" then sInput = "NORMAL"
        if rs(12) = "R" then sInput = "RECURRENTE"


    END IF 



        sinputK = ""
        sinputt = ""

 '   if  bModificar AND RS(3) > 10 AND RS(3) < 999 then
 '       sInput = "<font class=at><select name=" & rs(12).name & " onchange=submit()>"
 '       sInput = sInput & "<option value=U "
 '       if rs(12) = "U" then sInput = sInput & "selected"
 '       sInput = sInput & " >NORMAL<option value=R "
 '       if rs(12) = "R" then sInput = sInput & "selected"
 '       sInput = sInput & " >RECURRENTE</select></font>"
  '      sinputt = "<div id=AX" & rs(0) & " name=xxx style='display:none;' >" &  sInput & "</div>"

 '       sinputK = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
 '       sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
 '       sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
 '       sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
 '       sinputK = sinputK  & "none" & chr(34) & ";' "
 '       sinputK = sinputK  & "  href='javascript:void(0);' onclick='AX" & rs(0) & ".style.display=" & chr(34) & "block" & chr(34) & ";AY" & rs(0) & ".style.display=" & chr(34) & "none" & chr(34) & ";' "
 '       sinputK = sinputK  & "  TITLE='Modificar'  ><b>" & sImagenI & "</b></a>" 

'    end if


'        if rs(12) = "U" then sInput = "NORMAL"
'        if rs(12) = "R" then sInput = "RECURRENTE"

'        sInput = "<div id=AY" & rs(0) & " name=aaa style='display:block;' ><span>" &  sInput & "&nbsp;" &  sInputK & "</span></div>" &  sInputt


    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")


     sinput = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=WKFPASOSPROP.asp?WKFPSC=" & RS(2)
     sinput = sinput  & "  TITLE='Configuracion'  >" & sImagenI3 & "</a>" 


       celdas = celdas & ofv.GenCelda("",sc,"center","","","","",sInput,"si")


    filas = ofv.GenRow("","","","","","","",celdas,"si")



    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=WKFPASOS"
      formulario = formulario & ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario = formulario & ofv.GenForm("","","","","post","",filas,"si")
    end if


    posi = posi + 1
    rs.Movenext
  loop
  sHTML = sHTML  & ofv.GenTabla("","","","100%","","0","","-2","-1","","",formulario,"si")






      sInput = "<b>" & sImagenI & "&nbsp;Modificar</b>"
      celdas = ofv.GenCelda("","vt" ,"","15%","","","",sInput,"si")
      sInput = "<b>" & sImagenCP & "&nbsp;Modo&nbsp;del&nbsp;paso</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","15%","","","",sInput,"si")
      sInput = "<b>" & sImagenPI & "&nbsp;Paso&nbsp;inicial</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","15%","","","",sInput,"si")
      sInput = "<b>" & sImagenPF & "&nbsp;Paso&nbsp;final</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","15%","","","",sInput,"si")
      sInput = "<b>" & sImagenPC & "&nbsp;Paso&nbsp;del&nbsp;circuito</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","15%","","","",sInput,"si")
      sInput = "<b>" & sImagenI3 & "&nbsp;Configuracion</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","15%","","","",sInput,"si")


    formu = ofv.GenRow("","aaa height=30 VALIGN=MIDDLE ","","","","","",celdas,"si")
    formu = ofv.GenTabla("","aaa ","","100%","","0","","-2","-1","","",formu,"si")



  sHTML = sHTML & formu 

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "WKFPASOS.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  Buscar = ""
  
  sAccion = "WKFCIRCUITOS.asp"
  Celda = ofv.BotonVolver(sAccion)


 sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Celda,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

FUNCTION GENWKFPASO()

    WKFCOD = GTWKFPSC()

  StrSql = "Select MAX(ORDEN) From WKFPASOS WHERE WKFCOD = '" & session("WKF") & "' AND ORDEN < 999 "
  set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs3.eof then 
     NORDEN = RS3(0) + 10
  ELSE

     NORDEN = 20
  END IF

    strsql = "insert into WKFPASOS (WKFCOD,WKFPASCOD,ORDEN,DESCRIPCION,APR,RCH,TIPO) values ('" & session("WKF") & "','" & WKFCOD & "','" & NORDEN & "','NUEVO PASO WKF','1','1','U')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION

FUNCTION GENWKFPASOIF()

    WKFCOD = GTWKFPSC() & "I"
    NORDEN = 10


    strsql = "insert into WKFPASOS (WKFCOD,WKFPASCOD,ORDEN,DESCRIPCION,INI,TIPO) values ('" & session("WKF") & "','" & WKFCOD & "','" & NORDEN & "','INICIO','1','U')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

    session("WKFPSC") = WKFCOD
    GENWKFPSCTPC("INI")

    WKFCOD = GTWKFPSC() & "F"
    NORDEN = 999


    strsql = "insert into WKFPASOS (WKFCOD,WKFPASCOD,ORDEN,DESCRIPCION,FIN,TIPO) values ('" & session("WKF") & "','" & WKFCOD & "','" & NORDEN & "','FINAL','1','U')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

    session("WKFPSC") = WKFCOD
    GENWKFPSCTPC("FIN")

END FUNCTION

FUNCTION GENWKFPSCTPC(TIPO)

    strsql = "insert into WKFPSCTPC (WKFCOD,WKFPASCOD,TIPO,CAMEST,ESTADO,ANULA) values ('" & session("WKF") & "','" & session("WKFPSC") & "','" & TIPO & "','N','0','N')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION

%>