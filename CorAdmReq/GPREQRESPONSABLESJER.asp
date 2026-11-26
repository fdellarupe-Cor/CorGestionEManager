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

  SCODUSR = request.querystring("CODUSR")

  IF SCODUSR = "" THEN
     SCODUSR = SESSION("CODUSR")
  ELSE
     SESSION("CODUSR") = SCODUSR
  END IF



  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  Volver2 = replace(vuelta,"../","")



  set cn = ofv.conectar(ofv.strconn6)


  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


  StrSql = "Select * From GPREQRESPONSABLESJER WHERE CODUSR = '" & SCODUSR & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

    strsql = "insert into GPREQRESPONSABLESJER ("
    strsql = strsql & "CODUSR,USUARION1,USUARION2,SECTORN1,SECTORN2"
    strsql = strsql & ") values ("
    strsql = strsql & "'" & SCODUSR & "','NO','NO','0','0'"
    strsql = strsql & ")"
    
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

    StrSql = "Select * From GPREQRESPONSABLESJER WHERE CODUSR = '" & SCODUSR & "' "
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


  END IF





    sImagenC = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>&#0232;</FONT></B>"

    celdas = ""


  'Encabezados
  formulario = ""





  posi = 0
  IF Not rs.EOF THEN

    sform = "GPRESP" & rs(0)


       sc = "ut"

    FILAS = ""
    WKFSTYLER = " AAA "

         sInput = "RESP.NIVEL 1:"

         celdas = ""

         celdas = celdas & ofv.GenCelda("","UtI style='background-color:#888888;color:#ffffff;' ","","20%","","","",sInput,"si")

      sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
      sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
      sinput2 = sinput2  & "none" & chr(34) & ";'  href=agregarlista.asp?TABLA=GPREQRESPONSABLESJER&CodOBJETO=" &rs(0) & "&VOLVER=" & Volver2
      sinput2 = sinput2  & " TITLE='SUBTIPOS DE " & rs(1) & "'  ><b>" & sImagenC & "</b></a>" 

       if rs(2) <> "TODOS" THEN
          call getInfoUser(rs(2))  
       ELSE
          session("SGU_nombre")="SIN ASIGNAR"
       END IF
       
       if bModificar then
       
          sInput = session("SGU_nombre")   & sinput2
       else
          sInput = ofv.GenerarInput("",session("SGU_nombre"),"readonly","","60","dt","")
       end if

         celdas = celdas & ofv.GenCelda("","UtI  style='border-style:solid;border-width:1;BORDER-COLOR:#AAAAAA;' ","","80%","","","",sInput,"si")
         filas = filas & ofv.GenRow("",WKFSTYLER,"","","22","","",celdas,"si")






         sInput = "SECTOR.NIVEL 1:"

         celdas = ""

         celdas = celdas & ofv.GenCelda("","UtI style='background-color:#888888;color:#ffffff;' ","","","","","",sInput,"si")
                 set cn3 = ofv.conectar(ofv.strconn2)

                 XRESPO = "0" & RS(4)

                 if bModificar then
                    strsql = "Select * From ECO ORDER BY DESCRI"
                    sInput = ofv.GenerarCombo(sForm,RS(4).name,RS(4),strsql,cn3,0,1,"")
                 else
                    strsql = "Select * From ECO where ECO = " & XRESPO
                    set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
                    if not rsniv.eof then sniv = rsniv(1)
                    call ofv.cerrarconsulta(rsniv)
                    sInput = sniv
                 end if

                 call ofv.cerrarconn(cn3)

         celdas = celdas & ofv.GenCelda("","UtI  style='border-style:solid;border-width:1;BORDER-COLOR:#AAAAAA;' ","","","","","",sInput,"si")
         filas = filas & ofv.GenRow("",WKFSTYLER,"","","22","","",celdas,"si")










         sInput = "RESP.NIVEL 2:"

         celdas = ""

         celdas = celdas & ofv.GenCelda("","UtI style='background-color:#888888;color:#ffffff;' ","","","","","",sInput,"si")

      sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
      sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
      sinput2 = sinput2  & "none" & chr(34) & ";'  href=agregarlista.asp?TABLA=GPREQRESPONSABLESJER&CodOBJETO=" &rs(0) & "&VOLVER=" & Volver2
      sinput2 = sinput2  & " TITLE='SUBTIPOS DE " & rs(1) & "'  ><b>" & sImagenC & "</b></a>" 

       if rs(3) <> "TODOS" THEN
          call getInfoUser(rs(3))  
       ELSE
          session("SGU_nombre")="SIN ASIGNAR"
       END IF
       
       if bModificar then
       
          sInput = session("SGU_nombre")   & sinput2
       else
          sInput = ofv.GenerarInput("",session("SGU_nombre"),"readonly","","60","dt","")
       end if

         celdas = celdas & ofv.GenCelda("","UtI  style='border-style:solid;border-width:1;BORDER-COLOR:#AAAAAA;' ","","","","","",sInput,"si")
         filas = filas & ofv.GenRow("",WKFSTYLER,"","","22","","",celdas,"si")




         sInput = "SECTOR.NIVEL 2:"

         celdas = ""

         celdas = celdas & ofv.GenCelda("","UtI style='background-color:#888888;color:#ffffff;' ","","","","","",sInput,"si")
                 set cn3 = ofv.conectar(ofv.strconn2)

                 XRESPO = "0" & RS(5)

                 if bModificar then
                    strsql = "Select * From ECO ORDER BY DESCRI"
                    sInput = ofv.GenerarCombo(sForm,RS(5).name,RS(5),strsql,cn3,0,1,"")
                 else
                    strsql = "Select * From ECO where ECO = " & XRESPO
                    set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
                    if not rsniv.eof then sniv = rsniv(1)
                    call ofv.cerrarconsulta(rsniv)
                    sInput = sniv
                 end if

                 call ofv.cerrarconn(cn3)

         celdas = celdas & ofv.GenCelda("","UtI  style='border-style:solid;border-width:1;BORDER-COLOR:#AAAAAA;' ","","","","","",sInput,"si")
         filas = filas & ofv.GenRow("",WKFSTYLER,"","","22","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=GPREQRESPONSABLESJER"
      formulario = formulario & ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario = formulario & ofv.GenForm("","","","","post","",filas,"si")
    end if

    posi = posi + 1
  
  END IF

  sHTML = sHTML  & ofv.GenTabla("","","","80%","","0","","0","0","","",formulario,"si")


  'Botonera
  Agregar = ""

  Buscar = ""
  
  sAccion = "GPRESPONSABLES.asp?CantRegMover=" & CantRegMoverX
  Celda = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","",Celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if


response.write sHTML





%>