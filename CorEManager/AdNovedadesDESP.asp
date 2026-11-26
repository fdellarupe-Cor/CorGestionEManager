<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  CantRegMoverx = request.querystring("retorno")

  if CantRegMoverx = "" then  CantRegMoverx = request.querystring("Cantregmover")


  ID = request.querystring("ID")
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_ID=" & ID

  sHTML = ofv.FormHeader(session("FormN") & " - Envio del Mensaje")







  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  posi = 0




  Strsql = "Select * From novegenpar  where objid = " & ID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     manda = "N"

     if Clng(rs(6)) = 2 then

         Strsql = "Select * From novegenpardist  where codigo = " & ID
         set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if not rsx.eof then
            manda = "S"
         end if
         call ofv.cerrarconsulta(rsx)

     else
         manda = "S"

     end if





        select case clng(rs(5))
               case 0
                     sxopc4 = "ALTA"
               case 1
                     sxopc4 = "NORMAL"
               case 2
                     sxopc4 = "BAJA"

        end select

        select case clng(rs(6))

               case 1
                     sxopc5 = "TODOS"
               case 2
                     sxopc5 = "LISTA"

        end select



  'Encabezados
  Fila = ofv.Encabezados("Codigo","vigencia","dias","importancia","distribucion","contacto")



    Formulario = ""
    sForm = "form" & posi

    'Eliminar
    if   manda = "S" then
       sInput = "<font color=#008000 size=4><b>!</b></font>"
    else
       sInput = "<font color=#800000 size=4><b>x</b></font>"
    end if
      Celda = ofv.GenCelda("","cc","center","","","","",sInput,"si")


    'codigo
    sInput = "NGP" & cstr(rs(0))
    Celda = Celda & ofv.GenCelda("","cc","","5%","","","",sInput,"si")

    'Fecha Vigencia
      'sInput = cdate(clng(rs(1)) + clng(cdate("01-01-1900")))
 	sinput = Fechador(rs(1))
    Celda = Celda & ofv.GenCelda("","cc","","20%","","","","<b>" & sInput & "</b>","si")

    'dias
      sInput = rs(3)
    Celda = Celda & ofv.GenCelda("","cc","","20%","","","","<b>" & sInput & "</b>","si")


    'Importancia

                           select case clng(rs(5))
                                  case 0

                                        sxopc4x = "<font color=#800000 size=4><b>!</b></font>"
                                  case 1

                                        sxopc4x = "<font color=#008000 size=4><b>*</b></font>"
                                  case 2

                                        sxopc4x = "<font color=#000080 size=4><b>+</b></font>"

                            end select


      sInput = sxopc4x & "&nbsp;" & sxopc4
    Celda = Celda & ofv.GenCelda("","cc","","20%","","","","<b>" & sInput & "</b>","si")

    'distribucion
      sInput = sxopc5
    Celda = Celda & ofv.GenCelda("","cc","","20%","","","","<b>" & sInput & "</b>","si")

    'contacto
      sInput = rs(7)
    Celda = Celda & ofv.GenCelda("","cc","","20%","","","","<b>" & sInput & "</b>","si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")



    'Mensaje

    Celda = ofv.GenCelda("","tt colspan=7","center","50%","","","","Mensaje","si")
      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


    sInput = rs(4)

    Celda = ofv.GenCelda("","cc colspan=7","","50%","","","",sInput,"si")
      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


    Celda = ofv.GenCelda("","aaa colspan=7","center","50%","","","","&nbsp;","si")
      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")



    if   manda = "S" then


      sInput = ofv.GenerarInput("","CONFIRMA ENVIO","submit","","10","bt","")
      Celda = ofv.GenCelda("","aaa colspan=4","center","10%","","","",sInput,"si")

      sAccion = "AdnovedadesDespEnv.asp?CantRegMover=" & CantRegMoverx & "&ID=" & ID

      celda =  ofv.GenForm("Boton1" & rs(0),"",sAccion,"","post","",celda,"si")

      sInput = ofv.GenerarInput("","CANCELA ENVIO","submit","","10","bt","")
      Celda2 = ofv.GenCelda("","aaa colspan=3","center","10%","","","",sInput,"si")

      sAccion = "Adnovedades.asp?CantRegMover=" & CantRegMoverx

      Celda2 = ofv.GenForm("Boton2" & rs(0),"",sAccion,"","post","",Celda2,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda & Celda2,"si")

      Formulario =  Fila





    else


      sInput = "<font color=#800000 size=4><b>x</b></font><font  size=2>&nbsp;El mensaje <b>" & "NGP" & cstr(rs(0)) & "</b> no se puede enviar debido a que el tipo de"
      sInput = sInput & " distribucion establece una lista de usuarios, y no existe una para este mensaje.</font>"

      Celda = ofv.GenCelda("","cc colspan=7","","50%","","","",sInput,"si")
      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


    Celda = ofv.GenCelda("","aaa colspan=7","center","50%","","","","&nbsp;","si")
      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")




       sInput = ofv.GenerarInput("","VOLVER","submit","","10","bt","")
       Celda = ofv.GenCelda("","aaa colspan=7","center","10%","","","",sInput,"si")

       sAccion = "Adnovedades.asp?CantRegMover=" & CantRegMoverx


       celda =  ofv.GenForm("Boton1" & rs(0),"",sAccion,"","post","",celda,"si")


      Fila = fila & ofv.genrow("","","","","","","",Celda & Celda2,"si")

      Formulario =  Fila




    end if


  end if
  sHTML = sHTML & "<center><br>" & ofv.GenTabla("","","","80% ","","0","","0","1","","",Formulario,"si") & "</center>"




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

Function Fechador(wkfecha)


      wkfecha=clng(wkfecha)
      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))

      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function

%>