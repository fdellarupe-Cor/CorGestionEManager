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


  IF ID = "" THEN

     ID = SESSION("EMENCUESTA")

  ELSE

     SESSION("EMENCUESTA") = ID

  END IF

  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_ID=" & ID

  sHTML = ofv.FormHeader(session("FormN") & " - Propiedades del Mensaje")

  'Encabezados
  Formulario = ""





  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  posi = 0




  Strsql = "Select * From emencuestas  where objid = " & ID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

        sxopc3 = "<option value=1 >1 DIA</option> "
        sxopc3 = sxopc3 & "<option value=2 >2 DIAS</option> "
        sxopc3 = sxopc3 & "<option value=7 >7 DIAS</option> "
        sxopc3 = sxopc3 & "<option value=10 >10 DIAS</option> "
        sxopc3 = sxopc3 & "<option value=15 >15 DIAS</option> "
        sxopc3 = sxopc3 & "<option value=20 >20 DIAS</option> "
        sxopc3 = sxopc3 & "<option value=25 >25 DIAS</option> "
        sxopc3 = sxopc3 & "<option value=30 >30 DIAS</option> "
        sxopc3 = sxopc3 & "<option value=35 >35 DIAS</option> "
        sxopc3 = sxopc3 & "<option value=40 >40 DIAS</option> "

        sxopc3 = replace(sxopc3,"value=" & rs(3),"value=" & rs(3) & " selected")


        sxopc4 = "<option value=0 >ALTA</option> "
        sxopc4 = sxopc4 & "<option value=1 >NORMAL</option> "
        sxopc4 = sxopc4 & "<option value=2 >BAJA</option> "

        sxopc4 = replace(sxopc4,"value=" & rs(5),"value=" & rs(5) & " selected")

        sxopc5 = "<option value=1 >TODOS</option> "
        sxopc5 = sxopc5 & "<option value=2 >LISTA</option> "


        sxopc5 = replace(sxopc5,"value=" & rs(6),"value=" & rs(6) & " selected")

        VDOPCION = RS(19) & ""

        IF VDOPCION = "" THEN VDOPCION = "0"


        sxopcVD = "<option value=0 >NO</option> "
        sxopcVD = sxopcVD & "<option value=1 >IMAGEN</option> "
        sxopcVD = sxopcVD & "<option value=2 >VIDEO</option> "

        sxopcVD = replace(sxopcVD,"value=" & VDOPCION,"value=" & VDOPCION & " selected")


     fila = ""
     sForm = "formu"



      Celda =  ""
      Celda =  celda & ofv.GenCelda("","tt  ","","15%","","","","Codigo: ","si")
      sInput = "<b>NEC" & rs(0) & "</b>"
      Celda = celda & ofv.GenCelda("","cc style='border-style:solid;border-width:1;border-color:#888888;' ","","20%","","","",sInput,"si")

      Celda =  celda & ofv.GenCelda("","tt  ","","15%","","","","Tipo: ","si")

      scomK = "1" & chr(9) & "2"
      scomD = "SIMPLE" & chr(9) & "AVANZADA"

      sInput = ofv.gencomboFX(sForm,rs(2).name,rs(2),scomK,scomD," ")
      Celda = celda & ofv.GenCelda("","uti style='border-style:solid;border-width:1;border-color:#888888;' ","","20%","","","",sInput,"si")

      Celda =  celda & ofv.GenCelda("","tt  ","","15%","","","","Relacion: ","si")

      scomK = "1" & chr(9) & "2"
      scomD = "INDEPENDIENTE" & chr(9) & "MENSAJE DE NOTICIA"

      sInput = ofv.gencomboFX(sForm,rs(9).name,rs(9),scomK,scomD," ")
      Celda = celda & ofv.GenCelda("","uti style='border-style:solid;border-width:1;border-color:#888888;' ","","20%","","","",sInput,"si")


      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")






      Celda = ofv.GenCelda("","tt  ","","","","","","Fecha Vigencia: ","si")

      sInputD = ofv.GenerarInput(rs(3).name,rs(3).value,"HIDDEN",sForm,"1","dt","")

    if clng(rs(3)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSh(rs(3))
    end if

    sInput = dFechan & "&nbsp;" & sInputD

    sInput = "<u>" & sinput & "</u>"

    Celda = Celda & ofv.gencelda("","VTI STYLE='cursor:pointer;border-style:solid;border-width:1;border-color:#888888;' " & GTCSETCALR(sForm,rs(3).name,GTCSMT,GTCFCH,GTCFCHDAT,"",0,2,0,0,0),"","","","","",sInput,"si")


       Celda = Celda & ofv.GenCelda("","tt  ","","","","","","Fecha Expiracion: ","si")
      sInputD = ofv.GenerarInput(rs(4).name,rs(4).value,"HIDDEN",sForm,"1","dt","")

    if clng(rs(4)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSh(rs(4))
    end if

    sInput = dFechan & "&nbsp;" & sInputD

    sInput = "<u>" & sinput & "</u>"

    Celda = Celda & ofv.gencelda("","VTI STYLE='cursor:pointer;border-style:solid;border-width:1;border-color:#888888;' " & GTCSETCALR(sForm,rs(4).name,GTCSMT,GTCFCH,GTCFCHDAT,"",0,2,0,0,0),"","","","","",sInput,"si")


       Celda = Celda & ofv.GenCelda("","tt  ","","","","","","Fecha vista: ","si")
      sInputD = ofv.GenerarInput(rs(5).name,rs(5).value,"HIDDEN",sForm,"1","dt","")

    if clng(rs(5)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSh(rs(5))
    end if

    sInput = dFechan & "&nbsp;" & sInputD

    sInput = "<u>" & sinput & "</u>"

    Celda = Celda & ofv.gencelda("","VTI STYLE='cursor:pointer;border-style:solid;border-width:1;border-color:#888888;' " & GTCSETCALR(sForm,rs(5).name,GTCSMT,GTCFCH,GTCFCHDAT,"",0,2,0,0,0),"","","","","",sInput,"si")


      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")






      Celda =  ofv.GenCelda("","tt colspan=6 ","center","","","","","titulo","si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      sInput = "<textarea name=" & rs(17).name  & " onkeypress=" & chr(34)
      sInput = sInput & " if (this.value.length > 190) {alert('El texto es demasiado"
      sInput = sInput & " extenso'); return false;}" & chr(34)
   ''   sInput = sInput & " onblur=" & chr(34) & "document." & sform & ".submit();" & chr(34)
      sInput = sInput & " cols=120 rows=2>" & REPLACE(rs(17) & "","''","'") & "</textarea>"

      Celda =  ofv.GenCelda("","ut  colspan=6 style='border-style:solid;border-width:1;border-color:#888888;' ","center","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      Celda =  ofv.GenCelda("","tt colspan=6 ","center","","","","","copete","si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      sInput = "<textarea name=" & rs(18).name  & " onkeypress=" & chr(34)
      sInput = sInput & " if (this.value.length > 490) {alert('El texto es demasiado"
      sInput = sInput & " extenso'); return false;}" & chr(34)
   ''   sInput = sInput & " onblur=" & chr(34) & "document." & sform & ".submit();" & chr(34)
      sInput = sInput & " cols=120 rows=3>" & REPLACE(rs(18) & "","''","'") & "</textarea>"

      Celda =  ofv.GenCelda("","ut  colspan=6 style='border-style:solid;border-width:1;border-color:#888888;' ","center","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")




      Celda =  ofv.GenCelda("","tt colspan=6 ","center","","","","","Mensaje","si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      sInput = "<textarea name=" & rs(6).name  & " onkeypress=" & chr(34)
      sInput = sInput & " if (this.value.length > 950) {alert('El texto es demasiado"
      sInput = sInput & " extenso'); return false;}" & chr(34)
   ''   sInput = sInput & " onblur=" & chr(34) & "document." & sform & ".submit();" & chr(34)
      sInput = sInput & " cols=120 rows=6>" & REPLACE(rs(6) & "","''","'") & "</textarea>"

      Celda =  ofv.GenCelda("","ut  colspan=6 style='border-style:solid;border-width:1;border-color:#888888;' ","center","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")



      Celda =  ofv.GenCelda("","tt  ","","","","","","IMAGENES/VIDEO: ","si")
      sInput = "<select class=at name=" & rs(19).name  & selacc &  " >" & sxopcVD & "</select>"
      Celda = celda & ofv.GenCelda("","uti style='border-style:solid;border-width:1;border-color:#888888;' ","","","","","",sInput,"si")


   IF (RS(19) & "") = "1" THEN

      sInput = ofv.GenerarInput("","SELECCIONAR","BUTTON","","10","btSB","document.Botonx.submit();")
      Celda = celda & ofv.GenCelda("","uti colspan=4 style='border-style:solid;border-width:1;border-color:#888888;' ","","","","","",sInput,"si")

       sAccion = "Adencuestasimg.asp?Retorno=" & CantRegMoverx & "&ID=" & id
       FormuX = ofv.GenForm("Botonx","",sAccion,"","post","","&nbsp;","si")



   else

       Celda = celda &  ofv.GenCelda("","ut colspan=4 style='border-style:solid;border-width:1;border-color:#888888;' ","","","","","","&nbsp;","si")

   end if

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")




   IF (RS(19) & "") = "2" THEN

      Celda =  ofv.GenCelda("","tt colspan=6 ","center","","","","","URL DEL VIDEO:","si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      sInput = "<textarea name=" & rs(21).name  & " onkeypress=" & chr(34)
      sInput = sInput & " if (this.value.length > 490) {alert('El texto es demasiado"
      sInput = sInput & " extenso'); return false;}" & chr(34)
   ''   sInput = sInput & " onblur=" & chr(34) & "document." & sform & ".submit();" & chr(34)
      sInput = sInput & " cols=120 rows=3>" & REPLACE(rs(21) & "","''","'") & "</textarea>"

      Celda =  ofv.GenCelda("","ut  colspan=6 style='border-style:solid;border-width:1;border-color:#888888;' ","center","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

   END IF





      Celda =  ofv.GenCelda("","tt  ","","","","","","Importancia: ","si")

      sInput = "<select class=at name=" & rs(7).name  & selacc &  " >" & sxopc4 & "</select>"

 '     sInput = ofv.GenerarInput(rs(7).name,rs(7),"text",sForm,"60","dt","")
      Celda = celda & ofv.GenCelda("","uti style='border-style:solid;border-width:1;border-color:#888888;' ","","","","","",sInput,"si")

      Celda = celda &  ofv.GenCelda("","tt  ","","","","","","Tipo de Distribucion: ","si")

      sInput = "<select class=at name=" & rs(8).name  & selacc &  " >" & sxopc5 & "</select>"

'      sInput = ofv.GenerarInput(rs(8).name,rs(8),"text",sForm,"60","dt","")
      Celda = celda & ofv.GenCelda("","uti style='border-style:solid;border-width:1;border-color:#888888;' ","","","","","",sInput,"si")

      Celda = celda &  ofv.GenCelda("","tt  ","","","","","","Contacto: ","si")

      sInput = "<select class=at name=" & rs(8).name  & selacc &  " >" & sxopc5 & "</select>"

      sInput = "<b>" & rs(13) & "</b>"
      Celda = celda & ofv.GenCelda("","uti style='border-style:solid;border-width:1;border-color:#888888;' ","","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")



    sAccion = ""
    if bModificar then

      sInput1 = ofv.GenerarInput("","Estructura","button","","10","bt"," document.ESTR.submit();")
      sInput2 = ofv.GenerarInput("","Actualizar","submit","","10","bt","")

      Celdax = ""
      Celdax = celdax & ofv.GenCelda("","ut ","","","","","",sInput1,"si")
      Celdax = celdax & ofv.GenCelda("","ut ","","","","","",sInput2,"si")
      sAccion = "AdencuestasESTR.asp?CantRegMover=" & CantRegMoverx
      Formules =  ofv.genform("ESTR","",sAccion,"","post","","","si")

      Filax  = ofv.genrow("","","","","","","",Celdax,"si")
      sInput = ofv.GenTabla("","","","","","0","","0","1","","",Filax,"si")


      Celda = ofv.GenCelda("","ut colspan=6 style='background-color:#ddddff;border-style:solid;border-width:1;border-color:#888888;' ","right","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=emencuestas&volver=" & vuelta

    end if
    Formulario =  ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")








    posi = posi + 1





  end if
  sHTML = sHTML & "<center><br>" & ofv.GenTabla("","","","80% ","","0","","0","1","","",Formulario,"si") & "</center>" & FormuX & Formules




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  Agregar = ""
  buscar = ""

  sAccion = "Adencuestas.asp?CantRegMover=" & CantRegMoverx
  Celda = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no",Agregar,Buscar,Celda)

  sHTML = sHTML & GTCCALENDAR("M")


  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>