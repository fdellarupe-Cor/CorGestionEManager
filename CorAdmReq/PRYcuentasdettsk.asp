<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0

  sfuncion=" onload=" & chr(34) & "window.document.title='Cor OnTime - Consulta de Tareas';" & chr(34) & " "


  sHTML = ofv.menuHeader("",sfuncion)



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


if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if


    sopcion = request.querystring("opcion")

    cta = request.querystring("cta")

    sopcion = "VP"

    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"



''''''''''''''''''''''''''''''''''''''''''''''''

    sHTML = sHTML & "<div id='DX2' style='display:BLOCK;"
    sHTML = sHTML & "border-style:solid;border-width:1;border-color:#888888;"
    sHTML = sHTML & "background-color:" & session("APLBGD") & ";width=95%;"
    sHTML = sHTML & "' >"



      celdas = ofv.GenCelda("","TTT colspan=5 style='border-style:solid;border-width:1;border-color:#888888;' ","center","","","","","REGISTRO DE imputaciones DE la cuenta","si")

      filas = ofv.GenRow("","","","","10","","",celdas,"si")

          StrSql = "Select * From PROYTAREASADV where CODINPUT = '" & cta & "'  ORDER BY fupd DESC "
          set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rsX.eof then

                xTrColor = "#FFFFFF"







                celdas = ""
                celdas = celdas & ofv.GenCelda("","TTr","center","","","","","TAREA","si")
                celdas = celdas & ofv.GenCelda("","TTr","center","","","","","FECHA","si")
                celdas = celdas & ofv.GenCelda("","TTr","center","","","","","IMPUTADO","si")
                celdas = celdas & ofv.GenCelda("","TTr","center","","","","","PORCENTAJE","si")
                celdas = celdas & ofv.GenCelda("","TTr","center","","","","","RESPONSABLE","si")
                filas  = filas & ofv.GenRow("","","","","20","","",celdas,"si")


             do while not rsx.eof

                celdas = ""

''''''''''''''''''''''''''

                 XTSKDESCRI = ""

                 StrSql = "Select * From PROYtareas where TAREA = '" & rsX(2) & "' "
                 set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                 if not  rs1.eof then

                    XTSKDESCRI = RS1(4) & "." & rs1(5)
         


                 END IF

                celdas = celdas & ofv.GenCelda("","UT ","","30%","","","",XTSKDESCRI,"si")



''''''''''''''''''''''''''



                dFecha = cdate(rsX(3) + clng(cdate("01-01-1900")))
                SINPUTn1 = dFecha

                if rsX(7) < 4 then sInput = " H"


                if rsX(7) = 4 then sInput = " m"

                celdas = celdas & ofv.GenCelda("","uT","","10%","","","",SINPUTn1,"si")
                celdas = celdas & ofv.GenCelda("","ut","right","10%","","","",rsX(4) & sInput & "&nbsp;","si")
                celdas = celdas & ofv.GenCelda("","ut","right","10%","","","",rsX(5) & "%" & "&nbsp;&nbsp;&nbsp;","si")


                 sInput = "<B><I>SIN AFECTAR</I></B>"

                 strsql = "Select * From USUARIOS where ID = '" & rsX(6) & "' "
                 set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                 if not rsniv.eof then sniv = rsniv(1)
                 call ofv.cerrarconsulta(rsniv)
                 sInput = "<B>" & sniv & "</B>"


                celdas = celdas & ofv.GenCelda("","uti","","40%","","","",sInput,"si")

                filas = filas & ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","20","","",celdas,"si")
    


	        if xTrColor = "#FFFFFF" then
	           xTrColor = "#F4F4F4" 
	        ELSE
	           xTrColor = "#FFFFFF" 
	        END IF




             rsx.movenext
             loop  

          else

             celdas = ofv.GenCelda("","uT colspan=4  ","center","","","","","SIN IMPUTACIONES A LA FECHA","si")
             filas  = filas & ofv.GenRow("","","","","30","","",celdas,"si")



          end if

          sHTML = sHTML & ofv.GenTabla("","aa style='border-style:solid;border-width:1;border-color:#dddddd;' ","","98%","","0","","-1","-2","","",filas,"si")
          sHTML = sHTML & "<br>"





    sHTML = sHTML & "</div>"

 


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML



%>