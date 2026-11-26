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

  set cn = ofv.conectar(ofv.strconn5)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

 sopcion = request.querystring("opcion")
    session("CKT") = ""

 if sopcion = "SI" then

    WVERSION = request.querystring("VERSION")

    GENWKF(WVERSION)

 end if


 if sopcion = "DEL" then

        WKFCOD = request.querystring("WKFCOD")

        strsql = "DELETE FROM WKFCIRCUITO WHERE WKFCOD = '" & WKFCOD & "' "
  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


        strsql = "DELETE FROM WKFPASOSNV WHERE WKFCOD = '" & WKFCOD & "' "
  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


        strsql = "DELETE FROM WKFPASOSPRTYNV WHERE WKFCOD = '" & WKFCOD & "' "
  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


        strsql = "DELETE FROM WKFPSCRESPNV WHERE WKFCOD = '" & WKFCOD & "' "
  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)



 end if


    celdas = ""


   Celda = ""

   wkproccod = "&nbsp;"
    
   Celda = Celda & ofv.GenCelda("","uti  ","","90%","","","","LISTA DE CIRCUITOS","si")

   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")



  'Encabezados
  formulario =  ""

  ''''StrSql = "Select * From WKFCIRCUITO WHERE VERSION = 1 ORDER BY VERSION DESC, VALIDO, DESCRIPCION  "
  StrSql = "Select * From WKFCIRCUITO  ORDER BY VERSION DESC, VALIDO, DESCRIPCION  "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rs.eof then
     rs.move(CantRegMoverx)
     formulario = formulario & GENLISTA(RS)

  END IF




  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",formulario,"si")


  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "WKFCIRCUITOS.asp?Opcion=SI&VERSION=2"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  'Botonera
  VOLVER = ""

  Buscar = ""
    if bAgregar then
    sAccion = "WKFCIRCUITOS.asp?Opcion=SI&VERSION=1"
 '   Buscar = ofv.BotonAgregar(sAccion)
  end if



  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,VOLVER)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

FUNCTION GENWKF(XVERSION)

    IF XVERSION <> "2" THEN
       XVERSION = "1"
    END IF
       



    WKFCOD = GTWKF()

    strsql = "insert into WKFCIRCUITO (WKFCOD,DESCRIPCION,INICIO,CLASE,OBJTIPO,NIVELES,VERSION) values ('" & WKFCOD & "','NUEVO CIRCUITO WKF','A','M','REQ','3','" & XVERSION & "')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

    IF XVERSION = "2" THEN

         WKDESDE = 0

         for wkind=1 to 3

             WKDESDE = WKDESDE + 1

             WKFCODPRTY = GTWKFPSC() & "N" & WKDESDE
             WKRSPDESC  = "RESPONSABLE NIVEL " &  WKDESDE
             WKRSPTIPO  = ""

             IF WKDESDE = 1 THEN

                WKRSPDESC  = "SOLICITANTE"
                WKRSPTIPO  = "SOL"

             ELSEIF WKDESDE = 2 THEN

                WKRSPDESC  = "DESTINATARIO/RESPONSABLE"
                WKRSPTIPO  = "DES"

             END IF 



             strsql = "insert into WKFPSCRESPNV ("
             strsql = strsql & "WKFCOD,WKFPASRSPCOD,DESCRIPCION,NIVELRSP,"
             strsql = strsql & "MODORSP,ESTILORSP,SUPNIVELRSP,AUTOESCALARSP,"
             strsql = strsql & "TIPORSP,CODIGORSP"
             strsql = strsql & ") values ("
             strsql = strsql & "'" & WKFCOD & "','" & WKFCODPRTY & "','" & WKRSPDESC & "','" & WKDESDE & "',"
             strsql = strsql & "'U','A','8','S',"
             strsql = strsql & "'" & WKRSPTIPO & "',''"
             strsql = strsql & ")"
 '   response.write strsql & "<br>" & chr(13)
             call ofv.crearconsultaEx(StrSql,cn,1,parametros)


         next 

     end if



END FUNCTION


FUNCTION GENLISTA(RSX)

    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#F7AA22 SIZE=3>&#0143;</FONT></B>"
    sImagenI = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WebDINGS COLOR=#008000 SIZE=4>&#0194;</FONT></B>"
    sImagenI1 = "<B><FONT FACE=WEBDINGS COLOR=#000000 SIZE=3>&#0128;</FONT></B>"
    sImagenI2 = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=4>L</FONT></B>"
    sImagenI3 = "<B><FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=3>&#0232;</FONT></B>"
    sImagenW = "<B><FONT FACE=WEBDINGS COLOR=#6699CC SIZE=4>&#0194;</FONT></B>"
    sImagenW2 = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=4>&#0194;</FONT></B>"

    sImagenM1 = "<B><FONT FACE=WEBDINGS COLOR=#00AA00 SIZE=4>1</FONT></B>"
    sImagenM2 = "<B><FONT FACE=WEBDINGS COLOR=#005500 SIZE=4>&#0204;</FONT></B>"

    sImagenT1 = "<B><FONT FACE=WEBDINGS COLOR=#0000AA SIZE=4>&#0154;</FONT></B>"
    sImagenT2 = "<B><FONT FACE=WEBDINGS COLOR=#000055 SIZE=4>&#0156;</FONT></B>"

   xTrColor = "#FFFFFF"

  posi = 0
  prow = "N" 
  Do While (Not rsX.EOF) and (posi < CantRegMostrarx)

    sform = "REQ" & rsX(0)

    if prow = "S" then
       prow = "N"
       sr   = " bgcolor=#eeeeee "
    else 
       prow = "S"
       sr   = ""
    end if

       sr   = ""
       sc = "Ut"

       PERDEL = "N"


     StrSql = "Select * From tiporeqNV WHERE "
     StrSql = StrSql & " reqwkf1cod = '" & rsX(1) & "' or reqwkf2cod = '" & rsX(1) & "' or reqwkf3cod = '" & rsX(1) & "' or "
     StrSql = StrSql & " reqwkf4cod = '" & rsX(1) & "' or reqwkf5cod = '" & rsX(1) & "'  "
     set rsX5 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     IF RSX5.EOF THEN
        StrSql = "Select * From segcircuito WHERE "
        StrSql = StrSql & " (model = '" & rsX(1) & "' and fin = 'NO' )  "
        set rsX5 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        IF RSX5.EOF THEN
           PERDEL = "S"

        END IF


     END IF


     IF PERDEL = "S" THEN

        sInput2 = "<B><FONT COLOR='#FF0000'>X</FONT></B>"

        sinput = "<a  "
        sinput = sinput  & "  href='WKFcircuitos.asp?WKFCOD=" & RSx(1) & "&OPCION=DEL' "
        sinput = sinput  & "  TITLE='Eliminar Circuito " & rsx(2) & ".'  >" & sInput2 & "</a>" 
        sInput2 = sinput

        celdas = ofv.GenCelda("",sc & " STYLE='border-style:outset;border-color:#aa0000;' ","CENTER","2%","","","",sInput2,"si")


     else

        sInput2 = "-"
        celdas = ofv.GenCelda("",sc & " STYLE='border-style:solid;border-width:1;border-color:#eeeeee;' ","CENTER","2%","","","",sInput2,"si")

     END IF




       WKSPIMAGEN = sImagenW

       IF CSTR(RSX("VERSION") & "") = "2" THEN WKSPIMAGEN = sImagenW2





      celdas = celdas & ofv.GenCelda(""," Ut  ","","1%","","","",WKSPIMAGEN,"si")

       sInput = REPLACE(rsx(2) & ""," ","&nbsp;")

      celdas = celdas & ofv.GenCelda("",sc & "  style='border-right-style:solid;border-right-width:1;border-right-color:#888888;' ","","50%","","","",sInput,"si")


''' ALCANCE ''''''''''''''



        sniv = "S/D"
        strsql = "Select * From WKFTIPOS where TIPO = '" & rsx(6) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then sniv = rsniv(2)
        call ofv.cerrarconsulta(rsniv)
        sTwk = REPLACE(sniv & ""," ","&nbsp;")

      sInput = "" & sImagenO & ""
      celdas = celdas & ofv.GenCelda("",sc & " ","","1%","","","",sInput,"si")

      sInput = sTwk
      celdas = celdas & ofv.GenCelda("",sc & " " ,"","43%","","","",sInput,"si")



''' TIPO ''''''''''''''


        if rsX(4) = "A" then
           sInput = sImagenT1
           sTITLE = "AUTOMATICO"

        ELSEif rsX(4) = "S" then
           sInput = sImagenT2
           sTITLE = "A SOLICITUD"
        END IF

   '   celdas = celdas & ofv.GenCelda("",sc & " TITLE='" & STITLE & "' ","","1%","","","",sInput,"si")

''''''''''''''''''''''''''''''''


''' CLASE ''''''''''''''


        if rsX(5) = "E" then
           sInput = sImagenM1
           sTITLE = "EFECTIVO"

        ELSEif rsX(5) = "M" then
           sInput = sImagenM2
           sTITLE = "MODELO"
        END IF

    '  celdas = celdas & ofv.GenCelda("",sc & " TITLE='" & STITLE & "' ","","1%","","","",sInput,"si")

''''''''''''''''''''''''''''''''

 
    '  filas = ofv.GenRow("","aaa " & sr,"","","","","",celdas,"si")

     IF CSTR(RSX(11) & "") = "2" THEN

        sinput1 = "&nbsp;"

     ELSE

 


     sinput = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=WKFRESP.asp?WKF=" & RSx(1)
     sinput = sinput  & "  TITLE='Responsables'  >" & sImagenI1 & "</a>" 
     sinput1 = sinput

    END IF

      celdas = celdas & ofv.GenCelda("",sc & " " ,"","1%","","","",sInput1,"si")





     IF CSTR(RSX(11) & "") = "2" THEN

        WKASP = "WKFverNV.asp?WKF=" & RSx(1)

     ELSE

        WKASP = "WKFver.asp?WKF=" & RSx(1)

     END IF




     sinput = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "TARGET=_blank   href='" & WKASP & "' "
     sinput = sinput  & "  TITLE='Visualizacion'  >" & sImagenI2 & "</a>" 
     sinput2 = sinput

      celdas = celdas & ofv.GenCelda("",sc & " " ,"","1%","","","",sInput2,"si")


     IF CSTR(RSX(11) & "") = "2" THEN

        WKASP = "WKFPASOSNV.asp?WKF=" & RSx(1)

     ELSE

        WKASP = "WKFPASOS.asp?WKF=" & RSx(1)

     END IF



     sinput = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href='" & WKASP & "' "
     sinput = sinput  & "  TITLE='Configuracion'  >" & sImagenI3 & "</a>" 
     sinput3 = sinput

      celdas = celdas & ofv.GenCelda("",sc & " " ,"","1%","","","",sInput3,"si")



    '  sInput = sinput1 & "&nbsp;&nbsp;" & sinput2 & "&nbsp;&nbsp;" & sinput3
    '  celdas = celdas & ofv.GenCelda("","Ut colspan=2","","3%","","","",sInput,"si")





    filas = ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","","","",celdas,"si")


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF


    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=WKFCIRCUITO"
      formu = formu & ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formu = formu & ofv.GenForm("","","","","post","",filas,"si")
    end if


    posi = posi + 1
    rsX.Movenext
  loop


      sInput = "<b>" & sImagenI1 & "&nbsp;Responsables</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","15%","","","",sInput,"si")
      sInput = "<b>" & sImagenI2 & "&nbsp;Visualizacion</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","15%","","","",sInput,"si")
      sInput = "<b>" & sImagenI3 & "&nbsp;Configuracion</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","15%","","","",sInput,"si")


      sInput = "&nbsp;"
      celdas = ofv.GenCelda("","vt COLSPAN=9 style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#888888;' " ,"","1%","","","",sInput,"si")

 '   formu = formu & ofv.GenRow("","","","","","","",celdas,"si")


      sInput = sImagenW
      celdas = ofv.GenCelda("","vt" ,"","1%","","","",sInput,"si")
      sInput = "DESCRIPCION"
      celdas = celdas & ofv.GenCelda("","vt" ,"","","","","",sInput,"si")
      sInput = sImagenO
      celdas = celdas & ofv.GenCelda("","vt" ,"","1%","","","",sInput,"si")
      sInput = "ALCANCE"
      celdas = celdas & ofv.GenCelda("","vt" ,"","","","","",sInput,"si")


   ' formu = formu & ofv.GenRow("","","","","","","",celdas,"si")

      sInput = sImagenM1
      celdas = ofv.GenCelda("","vt" ,"","1%","","","",sInput,"si")
      sInput = "INICIO AUTOMATICO"
      celdas = celdas & ofv.GenCelda("","vt" ,"","","","","",sInput,"si")
      sInput = sImagenM2
      celdas = celdas & ofv.GenCelda("","vt" ,"","1%","","","",sInput,"si")
      sInput = "INICIO A SOLICITUD"
      celdas = celdas & ofv.GenCelda("","vt" ,"","","","","",sInput,"si")


 '   formu = formu & ofv.GenRow("","","","","","","",celdas,"si")

      sInput = sImagenT1
      celdas = ofv.GenCelda("","vt" ,"","1%","","","",sInput,"si")
      sInput = "MODELO"
      celdas = celdas & ofv.GenCelda("","vt" ,"","","","","",sInput,"si")
      sInput = sImagenT2
      celdas = celdas & ofv.GenCelda("","vt" ,"","1%","","","",sInput,"si")
      sInput = "EFECTIVO"
      celdas = celdas & ofv.GenCelda("","vt" ,"","","","","",sInput,"si")


  '  formu = formu & ofv.GenRow("","","","","","","",celdas,"si")


  GENLISTA = FORMU

END FUNCTION






%>