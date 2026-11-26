<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))

     ID = SESSION("EMENCUESTA")

  sblkid = request.querystring("BLKID")
  sitmid = request.querystring("ITMID")
  sOPCION = request.querystring("OPCION")

  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  


  sHTML = ofv.FormHeader(session("FormN"))



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


'''''''''''''''''''''''''''''''''''''''
'' INCORPORO BLOQUE A LA ENCUESTA

  if sOPCION = "SIB" then


     XXXX = FGENBLQ()


  end if

'''''''''''''''''''''''''''''''''''''''
'' INCORPORO ITEM AL BLOQUE DE LA ENCUESTA

  if sOPCION = "SIT" then


     XXXX = FGENITM(sblkid,"NUEVO ITEM")


  end if

'''''''''''''''''''''''''''''''''''''''





  Strsql = "Select * From emencuestas  where objid = " & ID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     fila = ""
     sForm = "formu"

      Celda =  ""
      Celda =  celda & ofv.GenCelda("","tt  ","","15%","","","","Codigo: ","si")
      sInput = "<b>NEC" & rs(0) & "</b>"
      Celda = celda & ofv.GenCelda("","cc style='border-style:solid;border-width:1;border-color:#888888;' ","","20%","","","",sInput,"si")

      Celda =  celda & ofv.GenCelda("","tt  ","","15%","","","","Tipo: ","si")

    'TIPO
    wkestado = rs(2) & ""
    WKTENC = rs(2) & ""
    select case wkestado

           case "1"

                 wkestadoD = "<font color='#0000AA' >SIMPLE</font>"

           case "2"

                 wkestadoD = "<font color='#993333' >AVANZADA</font>"


           CASE ELSE 


                 wkestadoD = "<font color='#000000' >SIN DETERMINAR</font>"


    end select

    sInput = wkestadoD

      Celda = celda & ofv.GenCelda("","uti style='border-style:solid;border-width:1;border-color:#888888;' ","","20%","","","",sInput,"si")

      Celda =  celda & ofv.GenCelda("","tt  ","","15%","","","","Relacion: ","si")


   'RELACION
    wkestado = rs(9) & ""

    select case wkestado

           case "1"

                 wkestadoD = "<font color='#336699' >INDEPENDIENTE</font>"

           case "2"

                 wkestadoD = "<font color='#993333' >MENSAJE DE NOTICIA</font>"


           case "3"

                 wkestadoD = "<font color='#009900' >DOCUEMENTO</font>"


           case "4"

                 wkestadoD = "<font color='#000099' >FORO</font>"

           CASE ELSE 


                 wkestadoD = "<font color='#000000' >SIN DETERMINAR</font>"


    end select

    sInput = wkestadoD

       Celda = celda & ofv.GenCelda("","uti style='border-style:solid;border-width:1;border-color:#888888;' ","","20%","","","",sInput,"si")


      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")






      Celda = ofv.GenCelda("","tt  ","","","","","","Fecha Vigencia: ","si")

 
    if clng(rs(3)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSh(rs(3))
    end if

    sInput = dFechan & "&nbsp;"

    Celda = Celda & ofv.gencelda("","UTI STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","","","","",sInput,"si")


       Celda = Celda & ofv.GenCelda("","tt  ","","","","","","Fecha Expiracion: ","si")
  
    if clng(rs(4)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSh(rs(4))
    end if

    sInput = dFechan & "&nbsp;"

    Celda = Celda & ofv.gencelda("","UTI STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","","","","",sInput,"si")

       Celda = Celda & ofv.GenCelda("","tt  ","","","","","","Fecha vista: ","si")

  
    if clng(rs(5)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSh(rs(5))
    end if

    sInput = dFechan & "&nbsp;"

    Celda = Celda & ofv.gencelda("","UTI STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","","","","",sInput,"si")


      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")






      Celda =  ofv.GenCelda("","tt colspan=6 ","center","","","","","titulo","si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      sInput = rs(17) & ""
      Celda =  ofv.GenCelda("","ut  colspan=6 style='border-style:solid;border-width:1;border-color:#888888;' ","center","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
      Formulario =  Fila
  
      posi = posi + 1

  end if
  sHTML = sHTML & "" & ofv.GenTabla("","","","80% ","","0","","0","1","","",Formulario,"si") & "<BR>"

'''''''''''' BLOQUES DE LA ENCUESTA ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''








  Strsql = "Select * From EMENCUESTASBLK WHERE ENCcodigo = 0" & ID
  Strsql = Strsql & " Order By BLKSECUEN,objid DESC"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if rs.eof then 
     XXXX = FGENBLQ()

     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  ELSE

     

  END IF




    Celda = ""

    Celda = Celda & ofv.gencelda("","VTI COLSPAN=2 STYLE='background-color:#bbbbbb;color:#ffffff;border-style:solid;border-width:1;border-color:#888888;' ","","5%","","","","SUBTITULO","si")
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;border-style:solid;border-width:1;border-color:#888888;'","","50%","","","","DESCRIPCION","si")
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;border-style:solid;border-width:1;border-color:#888888;'","","35%","","","","TIPO VOTO","si")
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;border-style:solid;border-width:1;border-color:#888888;'","","5%","","","","&nbsp;","si")

    Formulario = ofv.genrow("","","","","20","","",Celda,"si")

  posI = 0
    FormulES = ""
  Do While (Not rs.EOF)
    sForm = "formBLK" & posi
    sForm2 = "formBLKAD" & posi

    sblkcodigo = rs(0)
    sblktvoto  = rs(4)

''''''''''''''''''' items del bloque de encuesta ''''''''''''''''''''''''''''''''''''''''''''''''

       Strsql2 = "Select * From EMENCUESTASITM WHERE ENCcodigo = 0" & ID
       Strsql2 = Strsql2 & " and BLKcodigo = 0" & sblkcodigo
       Strsql2 = Strsql2 & " Order By ITMSECUEN,objid DESC"

       set rs2 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)

       if rs2.eof then 

          IF sblktvoto = 2 THEN
             XXXX = FGENITM(sblkcodigo,"SI")
             XXXX = FGENITM(sblkcodigo,"NO")
             set rs2 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
          END IF

       ELSE

       END IF
       posI2 = 0

       if rs2.eof then 

          sInput = "-------- SIN ITEMS DEFINIDOS --------------"
          Celda = ofv.gencelda("","UT","","","","","",sInput,"si")
          Formulario2 = ofv.genrow("","","","","","","",Celda,"si")

       ELSE

       Celda = ""

       Celda = Celda & ofv.gencelda("","VTI COLSPAN=2 STYLE='background-color:#DDDDDD;border-style:solid;border-width:1;border-color:#888888;' ","","10%","","","","ITEM","si")
       Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#DDDDDD;border-style:solid;border-width:1;border-color:#888888;'","","60%","","","","DESCRIPCION","si")
       Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#DDDDDD;border-style:solid;border-width:1;border-color:#888888;'","","30%","","","","VALOR","si")

       Formulario2 = ofv.genrow("","","","","20","","",Celda,"si")


       Do While (Not rs2.EOF)
          sForm3 = "form" & posi & "itm" & posi2

         'Eliminar
          if bSuprimir then
            sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs2(0)," ","_")
            sHref = sHref & "&Tabla=EMENCUESTASITM&volver=" & vuelta & "&Eliminar=NSNC"
            Celda = ofv.BotonEliminar(rs2(3),sHref)
          else
            Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
          end if

         'SECUENCIA
          sInput = ofv.GenerarInput(rs2(3).name,rs2(3),"text",sForm3,"3","dtN","")
          Celda = Celda & ofv.gencelda("","UT","","","","","",sInput,"si")
 
         'Descripcion
          sInput = ofv.GenerarInput(rs2(4).name,rs2(4),"text",sForm3,"70","dt","")
          Celda = Celda & ofv.gencelda("","UT","","","","","",sInput,"si")

          IF sblktvoto = 2 THEN
             sInput = "-----"
          ELSE
             sInput = ofv.GenerarInput(rs2(5).name,rs2(5),"text",sForm3,"10","dtN","")
          END IF
          Celda = Celda & ofv.gencelda("","UT","","","","","",sInput,"si")


          Fila = ofv.genrow("","","","","","","",Celda,"si")

          sAccion = ""
          if bModificar then
             sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs2(0)," ","*")
             sAccion = sAccion & "&tabla=EMENCUESTASITM&volver=" & vuelta
          end if
          Formulario2 = Formulario2 & ofv.genform(sForm3,"",sAccion,"","post","",Fila,"si")

          posi2 = posi2 + 1
          rs2.Movenext
      Loop

      END IF  

      sinputitm = ofv.GenTabla("","AAA style='border-style:solid;border-width:1;border-color:#888888;' BGCOLOR='#FFFFFF' ","","80% ","","0","","0","0","","",Formulario2,"si")



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Eliminar
    if bSuprimir AND posi2 = 0 then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=EMENCUESTASBLK&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(3),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'SECUENCIA
    sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"3","dtN","")
    Celda = Celda & ofv.gencelda("","UT style='background-color:#bbbbFF;border-style:solid;border-width:1;border-color:#888888;' ","","","","","",sInput,"si")

    'Descripcion
    sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"70","dt","")
    Celda = Celda & ofv.gencelda("","UT style='border-style:solid;border-width:1;border-color:#888888;' ","","","","","",sInput,"si")


    'TIPOVOTO
    scomK = "1" & chr(9) & "2"
    scomD = "SELECCION" & chr(9) & "SI/NO"
    sInput = ofv.gencomboFX(sForm,rs(4).name,rs(4),scomK,scomD," ")
    Celda = Celda & ofv.GenCelda("","UT style='border-style:solid;border-width:1;border-color:#888888;' " ,"","","","","",sInput,"si")


    IF (sblktvoto = 2 and posi2 < 2) or (sblktvoto = 1 and posi2 < 10 ) THEN
        sAccion = "AdencuestasESTR.asp?CantRegMover=" & CantRegMoverx & "&OPCION=SIT&BLKID=" & RS(0)
        FormulES = FormulES & ofv.genform(sForm2,"",sAccion,"","post","","","si")
        sInput1 = ofv.GenerarInput("","Agregar Item","button","","10","bt"," document." & sForm2 & ".submit();")
    ELSE
        sInput1 = "&nbsp;"
    END IF


    Celda = Celda & ofv.GenCelda("","UT style='border-style:solid;border-width:1;border-color:#888888;' " ,"","","","","",sInput1,"si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=EMENCUESTASBLK&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    Celda = ""
    sInput = sinputitm
    Celda = Celda & ofv.GenCelda("","UT COLSPAN=2 BGCOLOR='#FFFFFF' " ,"","","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","UT COLSPAN=3 BGCOLOR='#ddddFF' style='border-style:solid;border-width:1;border-color:#aaaaaa;' " ,"CENTER","","","","",sInput,"si")

    Fila = ofv.genrow("","AAA  ","","","","","",Celda,"si")
    Formulario = Formulario & Fila

    posi = posi + 1
    rs.Movenext
  Loop


 
    IF posi > 10 OR ( POSI > 1 AND WKTENC = 1)  THEN
       WKAGR = "N"
    END IF

    Celda = ""

    Celda = Celda & ofv.gencelda("","VTI COLSPAN=4 STYLE='background-color:#CCCCCC;color:#666666;border-style:solid;border-width:1;border-color:#888888;' ","","","","","","SUBTITULOS DE LA ENCUESTA","si")

    IF posi > 10 OR ( POSI > 1 AND WKTENC = "1")  THEN

       sAccion = "AdencuestasESTR.asp?CantRegMover=" & CantRegMoverx & "&OPCION=SIB"
       FormulADBLK = ofv.genform("FRMADDBLK","",sAccion,"","post","","","si")
       sInput1 = ofv.GenerarInput("","Agregar Subtitulo","button","","10","bt"," document.FRMADDBLK.submit();")

       Celda = Celda & ofv.gencelda("","UTI STYLE='background-color:#bbbbbb;color:#ffffff;border-style:solid;border-width:1;border-color:#888888;'","","","","","",sInput1,"si")
    ELSE
       FormulADBLK = ""
       Celda = Celda & ofv.gencelda("","UTI STYLE='background-color:#bbbbbb;color:#ffffff;border-style:solid;border-width:1;border-color:#888888;'","","","","","","&nbsp;","si")
    END IF

    Formulario = ofv.genrow("","","","","20","","",Celda,"si") & Formulario



  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si") & FormulES & FormulADBLK
  'Botonera
  Agregar = ""
  Buscar = ""

  sAccion = "ADENCUESTASPROP.asp?RETORNO=" & CantRegMoverx
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

FUNCTION FGENBLQ()

     Strsql1 = "Select BLKSECUEN From EMENCUESTASBLK WHERE ENCcodigo = 0" & ID
     Strsql1 = Strsql1 & " Order By BLKSECUEN DESC"
     set Wrs = ofv.crearconsultaEx(StrSql1,cn,1,parametros)

     if Wrs.eof then 
        SSECUEN = 10
     ELSE
        SSECUEN = WRS(0) + 10
     END IF


        strsql1 = "Insert into EMENCUESTASBLK "
        strsql1 = strsql1 & " (ENCcodigo,BLKSECUEN,BLKtitulo,BLKtipoVOTO, BLKTOTAL,"
        strsql1 = strsql1 & " BLKOBSERVACION "
        strsql1 = strsql1 & " ) Values ("
        strsql1 = strsql1 &  "'" & ID & "','" & SSECUEN & "','NUEVO SUBTITULO',"
        strsql1 = strsql1 &  "'1','0','NO'"
        strsql1 = strsql1 & ")"
  '''response.write strsql1 & "<br>"
        call ofv.crearconsultaEx(StrSql1,cn,1,volver)

    FGENBLQ = "OK"

END FUNCTION


FUNCTION FGENITM(WITMBLK,WITMDSC)

     Strsql1 = "Select ITMSECUEN From EMENCUESTASITM WHERE ENCcodigo = 0" & ID
     Strsql1 = Strsql1 & " AND BLKcodigo = 0" & WITMBLK
     Strsql1 = Strsql1 & " Order By ITMSECUEN DESC"
     set Wrs = ofv.crearconsultaEx(StrSql1,cn,1,parametros)

     if Wrs.eof then 
        SSECUEN = 10
     ELSE
        SSECUEN = WRS(0) + 10
     END IF


        strsql1 = "Insert into EMENCUESTASITM "
        strsql1 = strsql1 & " (ENCcodigo,BLKcodigo,ITMSECUEN,ITMtitulo,ITMVALOR, ITMTOTAL"
        strsql1 = strsql1 & " ) Values ("
        strsql1 = strsql1 &  "'" & ID & "','" & WITMBLK & "','" & SSECUEN & "','" & WITMDSC & "',"
        strsql1 = strsql1 &  "'0','0'"
        strsql1 = strsql1 & ")"
  '''response.write strsql1 & "<br>"
        call ofv.crearconsultaEx(StrSql1,cn,1,volver)

    FGENITM = "OK"

END FUNCTION

%>