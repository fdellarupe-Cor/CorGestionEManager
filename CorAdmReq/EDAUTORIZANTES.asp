<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
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

'  IF session("ModMain") <> "WF" THEN
'     bAgregar = FALSE
'     bModificar = FALSE
'     bSuprimir = FALSE
'  END IF

 sopcion = request.querystring("opcion")





  StrSql = "Select * From REQAUTOR where codigo = '" & session("NCliente") & "' ORDER BY FECHA "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)





    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","ttt colspan=3","","","","","",sImagen & "&nbsp;AUTORIZANTES","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","TT  ","center","","","","","AUTORIZANTE","si")
  celdas = celdas & ofv.GenCelda("","TT  ","center","","","","","SECTOR","si")
  celdas = celdas & ofv.GenCelda("","TT  ","center","","","","","FECHA","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  DO UNTIL RS.EOF


         sniv = "INEXISTENTE"
         snivS = RS(2)

      strsql = "Select * From USUARIOS where ID = '" & rs(3) & "' "
      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsniv.eof then
         sniv = rsniv(1)
         snivS = rsniv(21)
      END IF
      call ofv.cerrarconsulta(rsniv)
      sInput = "<B>" & sniv & "</B>"
     celdas = ofv.GenCelda("","CC","","","","","",sInput,"si")


     IF snivS > 0 THEN
        strsql = "Select * From ECO where ECO = " & SNIVS
        set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
        if not rsniv.eof then
           sniv = rsniv(1)
        END IF
        call ofv.cerrarconsulta(rsniv)
        sInput = sniv
        celdas = CELDAS & ofv.GenCelda("","CC","","","","","",sInput,"si")  
     ELSE
        celdas = CELDAS & ofv.GenCelda("","CC","","","","","","SIN AFECTAR","si")
     END IF 

        celdas = CELDAS & ofv.GenCelda("","CC","","","","","","SIN AFECTAR","si")


      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





    RS.MOVENEXT
    LOOP





    Celda = ofv.GenCelda("","abc colspan=2","center","100%","","","","&nbsp;","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")

    FilaS = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")

    Celda = ofv.GenCelda("","","center","","","","",FilaS,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")

'     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34) & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34) & "none" & chr(34) & ";' class=tt href=EDCONTACTOSBUS.asp?VASP=ES"
'     sinput = sinput  & " ><b>&nbsp;AGREGAR&nbsp;CONTACTO&nbsp;</b></a>" 
'     celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")
'     celdas = celdas & ofv.GenCelda("","","","","","","","&nbsp;","si")
'     fila = ofv.GenRow("","","","","","","",celdas,"si") 
'
'     Fila = ofv.GenTabla("","","","100%","","0","","0","0","","",fila,"si")

'    Celda = ofv.GenCelda("","","center","","","","",Fila,"si")
'    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",filas,"si")
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
  sHTML = sHTML  & ofv.GenTabla("","aa ","","98%","","0","","0","0","","",filas,"si")




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


''''''''''''''''''''''

  sHTML   = sHTML & CorPnlCOT()

  Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clpanel();' ", "S", "A", "N", "indFrm", "")

  Agregar = Celt1
  Buscar  = ""
  Celda   = ""

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

''''''''''''''''''''''


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML



FUNCTION GENCONTACTO(XUSR)




        STRSQL = "INSERT INTO REQCONTACTOS (CODIGO,CONTACTO) VALUES ("  
        STRSQL = STRSQL & "'" & session("NCliente") & "','" & XUSR & "')" 
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
''        response.write strsql











END FUNCTION

%>