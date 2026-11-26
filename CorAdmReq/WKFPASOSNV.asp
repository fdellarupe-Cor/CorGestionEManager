<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(session("FormN"))


  CantRegMostrarx = 30
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



 if sopcion = "DEL" then

        WKFPASCOD = request.querystring("WKFPASCOD")


        strsql = "DELETE FROM WKFPASOSNV WHERE WKFCOD = '" & session("WKF") & "' AND WKFPASCOD = '" & WKFPASCOD & "' "
  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


        strsql = "DELETE FROM WKFPASOSPRTYNV WHERE WKFCOD = '" & session("WKF") & "' AND WKFPASCOD = '" & WKFPASCOD & "' "
  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)





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

    IDWKF = RS(0)  

    SFORM = "Usuarios"

    sInput = "CIRCUITO DE WORKFLOW"
    celdas = ofv.GenCelda("","TTT","","20%","","","",sInput,"si")


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




    celdas = celdas & ofv.GenCelda("","UT","","","","","",sInput,"si")





    celdas = celdas & ofv.GenCelda("","UT","","","","","","&nbsp;NIVELES:","si")


    IF RS(10) = 0 THEN

       sInput = "<InPUT TYPE='BUTTON' CLASS='BT' VALUE='CONFIGURAR' ONCLICK='document.wkfresp.submit();' >"


    ELSE


      sInput = rs(10) & "&nbsp;<InPUT TYPE='BUTTON' CLASS='BT' VALUE='CONFIGURAR' ONCLICK='document.wkfresp.submit();' >"
       

    END IF




    celdas = celdas & ofv.GenCelda("","UT","","","","","",sInput,"si")

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(IDWKF," ","*") & "&volver=" & vuelta & "&Tabla=WKFCIRCUITO"
      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if

    filas = ofv.GenRow("","","","","","","",formulario0,"si")

  sHTML = sHTML  & ofv.GenTabla("","","","100% STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;' ","","0","","1","-1","","",filas,"si") & "<BR>"


    if bModificar then
      sAccion = "WKFPASOSRESPNV.asp?WKF=" & WKF
      formulario0 = ofv.GenForm("wkfresp","",sAccion,"","post","","","si")
    else
      formulario0 = ofv.GenForm("wkfresp","","","","post","","","si")
    end if

  sHTML = sHTML  & formulario0

  END IF



    celdas = ""



  'Encabezados
  formulario =  ofv.GenCelda("","ttR colspan=7","center","","","","","PASOS","si")
  formulario = ofv.GenRow("","","","","22","","",formulario,"si")


  StrSql = "Select * From WKFPASOSNV WHERE WKFCOD = '" & WKF & "' ORDER BY ORDEN"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF RS.EOF THEN
     GENWKFPASOIF()
     StrSql = "Select * From WKFPASOSNV WHERE WKFCOD = '" & WKF & "' ORDER BY ORDEN"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  END IF


  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    sform = "WKF" & rs(0)

    if rs(5) = "S" then
       sc = "CC"
    else
       sc = "TT"
    end if

       sc = "UT"


   PERDEL = "N"



   IF RS(6) = "S" OR RS(7) = "S" or rs(3) = 20 THEN
     sinput = "<B><FONT COLOR='#FF0000'>-</FONT></B>"

     PERDEL = "N"

   ELSE

     sinput = "<a style='text-decoration:none;color:#ffffff;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";'  href=PROYtask.asp?opcion=SBT&TSK="
     sinput = sinput  & " TITLE='SUBIR TAREA' ><b>" & sImagen1 & "</b></a>" 


     StrSql = "Select * From WKFPASOSPRTYNV WHERE PASOSIG = '" & RS(2) & "' "
     set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     IF RSX2.EOF THEN
        PERDEL = "S"


     else

        PERDEL = "N"

     END IF



   END IF



     IF PERDEL = "S" THEN

        sInput2 = "<B><FONT COLOR='#FF0000'>X</FONT></B>"

        sinputX = "<a  "
        sinputX = sinputX  & "  href='WKFPASOSNV.asp?WKFPASCOD=" & RS(2) & "&OPCION=DEL' "
        sinputX = sinputX  & "  TITLE='Eliminar PASO " & RS(4) & ".'  >" & sInput2 & "</a>" 
        sInput2 = sinputX

        celdas = ofv.GenCelda("",sc & " STYLE='border-style:outset;border-color:#aa0000;' ","CENTER","2%","","","",sInput2,"si")


     else

        sInput2 = "-"
        celdas = ofv.GenCelda("",sc & " STYLE='border-style:solid;border-width:1;border-color:#eeeeee;' ","CENTER","2%","","","",sInput2,"si")

     END IF




    celdas = celdas & ofv.GenCelda("","Ut  bgcolor=#6699cc ","CENTER","1%","","","",sinput,"si")






       sInput = rs(3) & ".&nbsp;"
       celdas = celdas & ofv.GenCelda("",sc,"right","","","","",sInput,"si")



       if bModificar then
          sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sform,"80","dt","")
       else
          sInput = ofv.GenerarInput("",rs(4),"readonly","","80","dt","")
       end if
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")


      '  if rs(6)  then     sInput = "" & sImagenPI & ""
      '  if rs(9)  then     sInput = "" & sImagenPF & ""
      '  if rs(7) OR RS(8) then     sInput = "" & sImagenPC & ""

     sInput = "" & sImagenPC & ""
  '  celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")



      sInput = "" & sImagenCP & ""
  '     celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

    if  XbModificar AND RS(3) > 10 AND RS(3) < 999 then


        scomK = "U" & chr(9) & "R"
        scomD = "NORMAL" & chr(9) & "RECURRENTE"

        if bModificar then
            sInput = ofv.gencomboFX(sForm,rs(20).name,rs(20),scomK,scomD," ")
        else
            sInput = ofv.GenerarInput("",rs(20),"readonly","","70","dt","")
        end if

    ELSE

        if rs(20) = "U" then sInput = "NORMAL"
        if rs(20) = "R" then sInput = "RECURRENTE"


    END IF 



        sinputK = ""
        sinputt = ""


    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")


     sinput = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=WKFPASOSPROPNV.asp?WKFPSC=" & RS(2)
     sinput = sinput  & "  TITLE='Configuracion'  >" & sImagenI3 & "</a>" 


       celdas = celdas & ofv.GenCelda("",sc,"center","","","","",sInput,"si")


    filas = ofv.GenRow("","","","","","","",celdas,"si")



    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=WKFPASOSNV"
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



 '' sHTML = sHTML & formu 

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "WKFPASOSNV.asp?Opcion=SI"
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

  StrSql = "Select MAX(ORDEN) From WKFPASOSNV WHERE WKFCOD = '" & session("WKF") & "' AND ORDEN < 999 "
  set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs3.eof then 
     NORDEN = RS3(0) + 10
  ELSE

     NORDEN = 20
  END IF

    strsql = "insert into WKFPASOSNV ("
    strsql = strsql & "WKFCOD,WKFPASCOD,ORDEN,DESCRIPCION,"
    strsql = strsql & "VALIDO,INI,FIN,SALTEAR,"
    strsql = strsql & "CANCELAR,NODO01,NODO02,NODO03,"
    strsql = strsql & "NODO04,NODO05,NODO06,NODO07,"
    strsql = strsql & "NODO08,NODO09,NODO10,TIPOCICLO,"
    strsql = strsql & "NIVELRSP,NROCICLO,TIEMPOESTIMADO,UNIDADTIEMPO,WKFACCION"
    strsql = strsql & ") values ("
    strsql = strsql & "'" & session("WKF") & "','" & WKFCOD & "','" & NORDEN & "','" & NORDEN & " - NUEVO PASO WKF',"
    strsql = strsql & "'S','N','N','N',"
    strsql = strsql & "'N','S','N','N',"
    strsql = strsql & "'N','N','N','N',"
    strsql = strsql & "'N','N','N','U',"
    strsql = strsql & "'0','1','1','1','D'"
    strsql = strsql & ")"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

    xxxx = GENWKFPSCTPC("3","N01")

END FUNCTION

FUNCTION GENWKFPASOIF()

    WKFCOD = GTWKFPSC() & "I"
    NORDEN = 10


    strsql = "insert into WKFPASOSNV ("
    strsql = strsql & "WKFCOD,WKFPASCOD,ORDEN,DESCRIPCION,"
    strsql = strsql & "VALIDO,INI,FIN,SALTEAR,"
    strsql = strsql & "CANCELAR,NODO01,NODO02,NODO03,"
    strsql = strsql & "NODO04,NODO05,NODO06,NODO07,"
    strsql = strsql & "NODO08,NODO09,NODO10,TIPOCICLO,"
    strsql = strsql & "NIVELRSP,NROCICLO,TIEMPOESTIMADO,UNIDADTIEMPO,WKFACCION"
    strsql = strsql & ") values ("
    strsql = strsql & "'" & session("WKF") & "','" & WKFCOD & "','" & NORDEN & "','INICIO',"
    strsql = strsql & "'S','S','N','N',"
    strsql = strsql & "'N','N','N','N',"
    strsql = strsql & "'N','N','N','N',"
    strsql = strsql & "'N','N','N','U',"
    strsql = strsql & "'1','1','1','1','IR'"
    strsql = strsql & ")"

    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

    session("WKFPSC") = WKFCOD
    xxxx = GENWKFPSCTPC("1","INI")

    WKFCOD = GTWKFPSC() & "F"
    NORDEN = 999

    strsql = "insert into WKFPASOSNV ("
    strsql = strsql & "WKFCOD,WKFPASCOD,ORDEN,DESCRIPCION,"
    strsql = strsql & "VALIDO,INI,FIN,SALTEAR,"
    strsql = strsql & "CANCELAR,NODO01,NODO02,NODO03,"
    strsql = strsql & "NODO04,NODO05,NODO06,NODO07,"
    strsql = strsql & "NODO08,NODO09,NODO10,TIPOCICLO,"
    strsql = strsql & "NIVELRSP,NROCICLO,TIEMPOESTIMADO,UNIDADTIEMPO,WKFACCION"
    strsql = strsql & ") values ("
    strsql = strsql & "'" & session("WKF") & "','" & WKFCOD & "','" & NORDEN & "','FINAL',"
    strsql = strsql & "'S','N','S','N',"
    strsql = strsql & "'N','N','N','N',"
    strsql = strsql & "'N','N','N','N',"
    strsql = strsql & "'N','N','N','U',"
    strsql = strsql & "'1','1','1','1','FR'"
    strsql = strsql & ")"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

    session("WKFPSC") = WKFCOD
    xxxx = GENWKFPSCTPC("2","FIN")

END FUNCTION

FUNCTION GENWKFPSCTPC(XTIPO,XITEMCOD)


    WKFCODPRTY = GTWKFPSC() & ITEMCOD

    IF XTIPO = "" THEN XTIPO = "3"
    IF XITEMCOD = "" THEN XITEMCOD = "N01"

    StrSql = "Select * From WKFPSCTIPOS WHERE TIPO = '" & XTIPO & "' "
    set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs3.eof then 
       XTIPODESC = RS3(2)
    ELSE

       XTIPODESC = "NO ENCONTRADO"
    END IF


    strsql = "insert into WKFPASOSPRTYNV ("
    strsql = strsql & "WKFCOD,WKFPASCOD,WKFPASPRTYCOD,DESCRIPCION,"
    strsql = strsql & "TIPOPSC,ITEMCOD,PASOSIG,PERMITECLC,"
    strsql = strsql & "CAMEST,ESTADO,WKFACCIONAD"
    strsql = strsql & ") values ("
    strsql = strsql & "'" & session("WKF") & "','" & session("WKFPSC") & "','" & WKFCODPRTY & "','" & XTIPODESC & "',"
    strsql = strsql & "'" & XTIPO & "','" & XITEMCOD & "','','S',"
    strsql = strsql & "'N','0','D'"
    strsql = strsql & ")"
 '   response.write strsql & "<br>" & chr(13)
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)





END FUNCTION

%>