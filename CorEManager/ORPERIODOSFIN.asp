<%@LCID = 11274%> 
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  sopcion  = request.querystring("opcion") & ""
  speriodo = request.querystring("periodo") & ""
  CantRegMostrarx = 10
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  sHTML = ofv.FormHeader(session("FORMN")) & "<BR><BR><BR>"

  'Encabezados
  Formulario = ""


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  WKSETTABLA = "ORPERIODOS"

  if sopcion = "SI" then

     sinifin = request.querystring("accion") & ""
     sfechai = clng(date()) - clng(cdate("01-01-1900"))

    if sinifin = "CL" or sinifin = "CLX" then

        Strsql = "update " & WKSETTABLA & " set ESTADO = 'CL', FINAL = '" & sfechai & "' "
        Strsql = Strsql & " where objid = " & speriodo
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)


      if sinifin = "CLX" then

        Strsql = "update orrskcircuito set ESTADO = 3, fecvigen = '" & sfechai & "' "
        Strsql = Strsql & " where periodo = " & speriodo & " and estado <> 3 and estado <> 4 "
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)

      end if


    end if

  response.redirect "ORPERIODOS.asp"


  end if


 DIM TABEST(10)

 FOR IEST=0 TO 10
     TABEST(IEST) = "N"
 NEXT


     sInput = "OBSERVACIONES"
     Celda = ofv.gencelda("","VTI COLSPAN=2 style='background-color:#336699;color:#ffffff;' ","","","","","",sInput,"si")

     Fila = ofv.genrow("","","","","25","","",Celda,"si")

     Formulario = Fila

     sInput = "&nbsp;"
     Celda = ofv.gencelda("","UT","","85%","","","",sInput,"si")

     sInput = "&nbsp;"
     Celda = Celda & ofv.gencelda("","UT","CENTER","15%","","","",sInput,"si")

     Fila = ofv.genrow("","","","","","","",Celda,"si")

     Formulario = Formulario & Fila


  sformobj = "rskfinal"


  Strsql = "Select DISTINCT ESTADO From ORRSKCIRCUITO WHERE APLIC = '" & WKAPLIC & "' AND "
  Strsql = Strsql & " ESTADO <> 4 AND ESTADO <> 3 "
  Strsql = Strsql & " Order By ESTADO "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then 


     sInput = "EL PERIODO ESTA EN CONDICIONES DE SER FINALIZADO. CONFIRMA SU FINALIZACION??"
     Celda = ofv.gencelda("","UTG  STYLE='border-style:SOLID;border-width:1;Border-color:#888888;'  ","","","","","",sInput,"si")

     sAccion = "ORPERIODOSFIN.asp?CantRegMover=" & CantRegMoverx & "&PERIODO=" & SPERIODO
     sAccion = sAccion & "&OPCION=SI&ACCION=CL"

     formobj1    = ofv.genform(sformobj,"",sAccion,"","post"," style='display:none;' ",sInput,"si")

     sInput2 = "<INPUT TYPE=BUTTON onclick=""disp_confirm('Confirma la finalizacion del periodo???','" & sformobj & "');"" VALUE='FINALIZAR'  STYLE='background-color:#990000;color:#ffffff;' >" 
     Celda = Celda & ofv.gencelda("","ut STYLE='border-style:SOLID;border-width:1;Border-color:#888888;' ","CENTER","","","","",sInput2,"si")

     Fila = ofv.genrow("","","","","","","",Celda,"si")
     Formulario = Formulario & Fila

     sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","AAA ","","60%","","0","","2","1","","",Formulario,"si") & "</CENTER>" & formobj1

  ELSE
  posi = 0

                    Strsql = "SELECT * From ORRSKCONFIG where APLIC = '" & wkaplic & "' "
                    set rsX3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                    if  not rsX3.eof then
                        XFSENWKFAUTOSET = rsX3(3)
                        XFSENWKFOBSSET  = rsX3(4)
                    end if     


  Do While (Not rs.EOF)

     TABEST(RS(0)) = "S"

    posi = posi + 1
    rs.Movenext
  Loop



    IF TABEST(2) = "S" OR TABEST(5) = "S" THEN


       sInput = "EL PERIODO NO ESTA EN CONDICIONES DE SER FINALIZADO, EXISTEN PROCESOS CON <BR>CIRCUITOS DE OBSERVACIONES O DE AUTORIZACION AUN PENDIENTES. "
       Celda = ofv.gencelda("","UTG STYLE='border-style:SOLID;border-width:1;Border-color:#888888;' COLSPAN=2 ","","","","","",sInput,"si")

 
       Fila = ofv.genrow("","","","","","","",Celda,"si")
       Formulario = Formulario & Fila


    ELSEIF TABEST(1) = "S" OR TABEST(6) = "S" THEN


      IF XFSENWKFAUTOSET = 2 OR (XFSENWKFOBSSET  = 2 AND TABEST(1) = "S" ) THEN

         sInput = "EL PERIODO NO ESTA EN CONDICIONES DE SER FINALIZADO EN SU SITUACION ACTUAL, EXISTEN PROCESOS AUN NO CONFIRMADOS. "
         Celda = ofv.gencelda("","UTG  STYLE='border-style:SOLID;border-width:1;Border-color:#888888;' COLSPAN=2  ","","","","","",sInput,"si")

         Fila = ofv.genrow("","","","","","","",Celda,"si")
         Formulario = Formulario & Fila

     ELSE 


         sInput = "EL PERIODO NO ESTA EN CONDICIONES DE SER FINALIZADO EN SU SITUACION ACTUAL, <BR>EXISTEN PROCESOS AUN NO CONFIRMADOS. "
         sInput = sInput & "<BR>DESEA CONFIRMARLOS EN FORMA MASIVA Y FINALIZAR EL PERIODO IGUALMENTE??. "
         Celda = ofv.gencelda("","UTG  STYLE='border-style:SOLID;border-width:1;Border-color:#888888;'  ","","","","","",sInput,"si")


         sAccion = "ORPERIODOSFIN.asp?CantRegMover=" & CantRegMoverx & "&PERIODO=" & SPERIODO
         sAccion = sAccion & "&OPCION=SI&ACCION=CLX"

         formobj1    = ofv.genform(sformobj,"",sAccion,"","post"," style='display:none;' ",sInput,"si")

         sInput2 = "<INPUT TYPE=BUTTON onclick=""disp_confirm('Confirma la finalizacion del periodo???','" & sformobj & "');"" VALUE='FINALIZAR'  STYLE='background-color:#990000;color:#ffffff;' >" 
         Celda = Celda & ofv.gencelda("","UT  STYLE='border-style:SOLID;border-width:1;Border-color:#888888;'  ","CENTER","","","","",sInput2,"si")

         Fila = ofv.genrow("","","","","","","",Celda,"si")
         Formulario = Formulario & Fila

      END IF


    END IF




         sInput = "DETALLE DE PROCESOS AUN NO CONFIRMADOS. "
         Celda = ofv.gencelda("","UTG  STYLE='border-style:SOLID;border-width:1;Border-color:#888888;'  ","","","","","",sInput,"si")


         sAccion = "ORLISTAPENDIENTES.asp?PERIODO=" & SPERIODO

         formobj1    = ofv.genform(sformobj,"",sAccion,"_BLANK","post"," style='display:none;' ",sInput,"si")

         sInput2 = "<INPUT TYPE=BUTTON onclick=""disp_confirm('Desea consultar la lista de procesos pendientes???','" & sformobj & "');"" VALUE='LISTA'  STYLE='background-color:#336699;color:#ffffff;' >" 
         Celda = Celda & ofv.gencelda("","UT  STYLE='border-style:SOLID;border-width:1;Border-color:#888888;'  ","CENTER","","","","",sInput2,"si")

         Fila = ofv.genrow("","","","","","","",Celda,"si")
         Formulario = Formulario & Fila



    sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","AAA ","","60%","","0","","2","1","","",Formulario,"si") & "</CENTER>" & formobj1

  END IF

  Agregar = ""
  Buscar  = ""

  sAccion = "ORPERIODOS.asp"
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no",Agregar,Buscar,celda)




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


    sHTML = sHTML & "<script type='text/javascript'>"
    sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
    sHTML = sHTML & "  {"
    sHTML = sHTML & "  var r=confirm(disp_msg);"
    sHTML = sHTML & "  if (r==true)"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
  '  sHTML = sHTML & "    alert(fr.name);"
    sHTML = sHTML & "    fr.submit();"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    alert('Se cancelo la accion!');"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  }"

    sHTML = sHTML & "</script>"

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>