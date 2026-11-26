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
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  sopcion  = request.querystring("opcion") & ""
  lote     = request.querystring("lote")
  sHTML = ofv.FormHeader(session("FormN"))


    APLCHECK()
 dim fs,fname
   set fname=nothing


    sInput = "GENERACION ARCHIVO ENVIO EVENTOS DE PERDIDA A BCRA"
    Celda = ofv.gencelda("","TT ","","90%","","","",sInput,"si")


    Fila = ofv.genrow("","","","","","","",Celda,"si")

  Formulario = Fila
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si") & "<BR>"

  FILA = ""

  LINEAP = ""





  'Encabezados
  Formulario = ofv.Encabezados("lote","DISEÑO","ENTIDAD","ESTADO","DESCRIPCION","FECHA","GENERAR","archivo&nbsp;txt")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)



  if sOPCION="GENERAR" then

   call genTxt(lote )

  end if



  WKSETTABLA = "ORREGINFOBCRA"

  Strsql = "Select * From " & WKSETTABLA & " where estado IN (2,3,4) ORDER BY ANO DESC, TRIMESTRE DESC "


  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi


    CELDA = ofv.gencelda("","UT  " ,"","0%","","","","&nbsp;","si")

    'LOTE
    sInput = RS(0)
    Celda = Celda & ofv.gencelda("","UT  STYLE='border-style:solid;border-width:1;border-color:#888888;' " ,"","5%","","","",sInput,"si")

    sInput = RS("CODDISENO")
    Celda = Celda & ofv.gencelda("","UT  STYLE='border-style:solid;border-width:1;border-color:#888888;'  " ,"","5%","","","",sInput,"si")

    sInput = RS("CODENTIDAD")
    Celda = Celda & ofv.gencelda("","UT  STYLE='border-style:solid;border-width:1;border-color:#888888;' " ,"","5%","","","",sInput,"si")




   'ESTADO


    select case RS("ESTADO")


           case "0"

                 wkestadoD = "<font color='#000000' >NUEVO</font>"

           case "1"

                 wkestadoD = "<font color='#993333' >REVERTIDO</font>"


           case "2"

                 wkestadoD = "<font color='#009900' >PROCESADO</font>"


           case "3"

                 wkestadoD = "<font color='#000099' >GENERADO</font>"

           case "4"

                 wkestadoD = "<font color='#336699' >ENVIADO</font>"

           CASE ELSE


                 wkestadoD = "<font color='#000000' >????</font>"


    end select

    sInput = "&nbsp;" & wkestadoD & "&nbsp;"
    Celda = Celda & ofv.gencelda("","VT STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","10%","","","",sInput,"si")

    sInput = rs("descri")
    Celda = Celda & ofv.gencelda("","UT STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","30%","","","",sInput,"si")



    SINPUT= "TRIMESTRE : " & RS("TRIMESTRE") & " AÑO :" & RS("ANO")

    Celda = Celda & ofv.gencelda("","UT STYLE='border-style:solid;border-width:1;border-color:#888888;'  ","","25%","","","",SINPUT,"si")

 IF RS("ESTADO") > "1" AND RS("ESTADO") < "4" THEN
	  sAccion = "ORGENARCHRUNOR.asp?OPCION=GENERAR&LOTE=" & RS("CODLOTE")
	  BTIT1 = "<INPUT TYPE='BUTTON' VALUE='GENERAR' CLASS='cc' title='Procesar el lote.' STYLE='cursor:pointer;background-color:#000099;color:#ffffff;' onclick=""disp_confirm('Confirma generar este lote???','FINI" & posi & "');""  > "
	  Formu1 =formu1 &  ofv.GenForm("FINI" & posi,"",sAccion,"","post","","","si")
	  Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","25%","","","",BTIT1,"si")
 ELSE
     Celda = Celda & ofv.gencelda("","UT STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","10%","","","","&nbsp;-","si")
 END If


 IF RS("ESTADO") > "2"  THEN

    if session("APLMDCLI")="BC" THEN

       if rs("archivo1")&"" <>"" then
          sInput1 = " <a  href='ORGENARCHRUNORLST.asp?OPCION=1&LOTE=" & RS("CODLOTE") & "&ARCHIVO=" & rs("archivo1") & "' target='_blank' ><font color=#ff0000>" & rs("archivo1") & "</font></a>"
          delimc="<br>"
       end if

       if rs("archivo2")&"" <>"" then
	  sInput1 = sinput1 & delimc &  " <a  href='ORGENARCHRUNORLST.asp?OPCION=2&LOTE=" & RS("CODLOTE") & "&ARCHIVO=" & rs("archivo2") & "' target='_blank' ><font color=#ff0000>&nbsp;" & rs("archivo2") & "</font></a>"
	  delimc="<br>"
       end if

       if rs("archivo3")&"" <> "" then
	  sInput1 = sinput1 & delimc &  " <a  href='ORGENARCHRUNORLST.asp?OPCION=3&LOTE=" & RS("CODLOTE") & "&ARCHIVO=" & rs("archivo3") & "' target='_blank' ><font color=#ff0000>&nbsp;" & rs("archivo3") & "</font></a>"
       end if


    ELSE


       if rs("archivo1")&"" <>"" then
          sInput1 = " <a  href=interfases/bcra/" & rs("archivo1") & " target='_blank' ><font color=#ff0000>" & rs("archivo1") & "</font></a>"
          delimc="<br>"
       end if

       if rs("archivo2")&"" <>"" then
          sInput1 = sinput1 & delimc &  " <a  href=interfases/bcra/" & rs("archivo2") & " target='_blank' ><font color=#ff0000>&nbsp;" & rs("archivo2") & "</font></a>"
	  delimc="<br>"
       end if


       if rs("archivo3")&"" <> "" then
	  sInput1 = sinput1 & delimc &  " <a  href=interfases/bcra/" & rs("archivo3") & " target='_blank' ><font color=#ff0000>&nbsp;" & rs("archivo3") & "</font></a>"
       end if

    END IF

 else
     sInput1 = "&nbsp;"
 end if

	  Celda = Celda & ofv.gencelda("","UT STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","10%","","","",sInput1,"si")


       Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")


'''''' BUSCA RECTIFICATORIAS

       wwwktri = "N"
       wwwkano = "N"


       Strsql = "Select * From ORREGINFOBCRAEXT where CODLOTE = '" & RS(1) & "' ORDER BY ANO, TRIMESTRE"

       set rsRR = ofv.crearconsultaEx(StrSql,cn,1,parametros)

       Do While (Not rsRR.EOF)

          if wwwktri <> (RSRR("TRIMESTRE") & "") or wwwkano <> (RSRR("ANO") & "") then


             if wwwktri <> "N" then


                CELDA = ofv.gencelda("","UT COLSPAN=5 " ,"","0%","","","","&nbsp;","si")
          
                sInput = "RECTIFICATIVA"
                Celda = Celda & ofv.gencelda("","UT STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","30%","","","",sInput,"si")

                SINPUT = "TRIMESTRE : " & wwwktri & " AÑO :" & wwwkano
                Celda = Celda & ofv.gencelda("","UT STYLE='border-style:solid;border-width:1;border-color:#888888;'  ","","25%","","","",SINPUT,"si")

                Celda = Celda & ofv.gencelda("","UT STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","10%","","","","&nbsp;-","si")



                Celda = Celda & ofv.gencelda("","UT STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","10%","","","",wwwkarch,"si")

                Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")

             end if

             wwwkarch = "" 

             wwwktri = RSRR("TRIMESTRE") & ""
             wwwkano = RSRR("ANO") & ""

          end if



                IF RS("ESTADO") > "2"  THEN
 
                   if session("APLMDCLI")="BC" THEN

                      if (rsRR("archivo") & "") <>"" then
                          wwwkarchOPC = "3R"
                          IF INSTR((rsRR("archivo") & ""),"RRIESGO") > 0 THEN wwwkarchOPC = "1R"
                          wwwkarch = wwwkarch & " <a  href='ORGENARCHRUNORLST.asp?OPCION=" & wwwkarchOPC & "&LOTE=" & RS("CODLOTE") & "&ARCHIVO=" & rsRR("archivo") & "' target='_blank' ><font color=#ff0000>" & rsRR("archivo") & "</font></a><br>"
                      else
                          wwwkarch = wwwkarch & "-&nbsp;<br>"
                      end if

                   ELSE

                      if (rsRR("archivo") & "") <>"" then
                          wwwkarch = wwwkarch & " <a  href=interfases/bcra/" & rsRR("archivo") & " target='_blank' ><font color=#ff0000>" & rsRR("archivo") & "</font></a><br>"
                      else
                          wwwkarch = wwwkarch & "-&nbsp;<br>"
                      end if

                   END IF

                else
                   wwwkarch = wwwkarch & "-&nbsp;<br>"
                end if





          rsRR.Movenext
       Loop


             if wwwktri <> "N" then


                CELDA = ofv.gencelda("","UT COLSPAN=5 " ,"","0%","","","","&nbsp;","si")
          
                sInput = "RECTIFICATIVA"
                Celda = Celda & ofv.gencelda("","UT STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","30%","","","",sInput,"si")

                SINPUT = "TRIMESTRE : " & wwwktri & " AÑO :" & wwwkano
                Celda = Celda & ofv.gencelda("","UT STYLE='border-style:solid;border-width:1;border-color:#888888;'  ","","25%","","","",SINPUT,"si")

                Celda = Celda & ofv.gencelda("","UT STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","10%","","","","&nbsp;-","si")



                Celda = Celda & ofv.gencelda("","UT STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","10%","","","",wwwkarch,"si")

                Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")

             end if


''''''''''''


    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",formulario & FILA & Formu1,"si")

  'Botonera

 ' str1 = "Select descri From " & WKSETTABLA  & " order by descri"  ' orden pagina
 ' str2 = "Select descri From " & WKSETTABLA   & " order by descri"    ' orden alfabetico
  Buscar =""' ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))


  Celda = ""

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

    sHTML = sHTML & "<script type='text/javascript'>"
    sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
    sHTML = sHTML & "  {"
    sHTML = sHTML & "  var r=confirm(disp_msg);"
    sHTML = sHTML & "  if (r==true)"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
  '''  sHTML = sHTML & "    alert(fr.name);"
    sHTML = sHTML & "    fr.submit();"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    alert('Se cancelo la accion!');"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  }"

''    sHTML = sHTML & "  }"
    sHTML = sHTML & "</script>"

  sHTML = sHTML & "</body></html>"



end if

''''''''''''''''''''''''''''''''''''''''''''''''''''''' FUNCION DE GENERACION DEL LOTE ''''''''''''''''''''''''

function genTxt(lote)

 dim strsql
 dim rsl

 strsqlx = "delete from ORREGINFOBCRAEXT where CODLOTE =  '" & LOTE & "' "
 ''response.write strsqlx
 SET xxxx = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)


 if session("APLMDCLI")="BC" THEN

    ruta = ""
    pos  = 0

       strsqlx = "delete from orrepositor where tlote =  '" & LOTE & "' "
      ''response.write strsqlx

       SET xxxx = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)

 else

    ruta = server.mappath(request.servervariables("script_name"))
    pos = instrrev(ruta, "\")
    ruta = MID(ruta, 1, pos)
    set fs=Server.CreateObject("Scripting.FileSystemObject")
 end if


 strsql="select * from orconfigbcra "
 SET RSconfig = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     GENERAIDPRES = RSconfig("GENERAIDPRES")
     CODIGOCUP    = RSconfig("CODIGOCUP")

 CALL OFV.CERRARCONSULTA(RSconfig)


 STRSQL=" SELECT A.OBJID,A.TRIMESTRE,A.ANO,A.codentidad ,A.coddiseno  FROM ORREGINFOBCRA A where a.codlote='"  & lote &"'"
'response.write strsql
 SET RSL=ofv.crearconsultaEx(StrSql,cn,1,parametros)

     codentidad = strpad(rsL("codentidad"),5,"0",2)
   'response.write codentidad
     coddiseno  = strpad(rsL("coddiseno"),4,"0",2)
     XX         = GetFecTrim(RSL("TRIMESTRE"),RSL("ANO"),FI,FF)
     FFINT      = FechadorCor(FF, "RI")

 CALL OFV.CERRARCONSULTA(RSL)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



 WANTAA = 0
 WANTTT = 0


 strsql = "SELECT eventoraiz, secuencia , monto, FECinicio , FECDESCUBRIMIENTO ,"
 strsql = strsql & " FECFINAL , FECREGCONTABLE, descripcion, TIPOEVENTO, LINEANEGOCIO,"
 strsql = strsql & " SUCURSAL, PROCESO, PRODUCTO, CARACTER, PARTIDACONTABLE,"
 strsql = strsql & " COBERPOLIZA, VINCOTROCRIESGO , AREASOPORTE, ESTADO, YANIO, TRIMESTRE"
 strsql = strsql & " FROM ORREGINFOBCRADET  WHERE CODLOTE='" &  LOTE & "' AND ESTADO = 'R'  "
 strsql = strsql & " order by  YANIO, TRIMESTRE, eventoraiz, secuencia"
'response.write strsql

 SET XRSL = ofv.crearconsultaEx(StrSql,cn,1,parametros)

 if not Xrsl.eof then


 do while not Xrsl.eof

    WACTAA = XrsL("YANIO")
    WACTTT = XrsL("TRIMESTRE")

    IF (WACTAA <> WANTAA) OR (WACTTT <> WANTTT) THEN

        IF WANTAA = 0 THEN


        ELSE

           if session("APLMDCLI")="BC" THEN

           else
              fname.Close
              set fname=nothing
           ''''   set fs=nothing
           end if

           Sid   = lote

           strsql = "Insert into ORREGINFOBCRAEXT "
           strsql = strsql & " (CODLOTE,ESTADO,TRIMESTRE,ANO,TOTAL,ARCHIVO"
           strsql = strsql & " ) Values ("
           strsql = strsql &  "'" & Sid & "','1','" & WANTTT & "','" & WANTAA & "','" & wwTOTRR & "','" & arnom1 & "'"
           strsql = strsql & ")"
        ''  response.write strsql & "<br>"
           call ofv.crearconsultaEx(StrSql,cn,1,volver)



        END IF

        wwTOTRR = 0

        WANTAA  = XrsL("YANIO")
        WANTTT  = XrsL("TRIMESTRE")

        XXXXX      = GetFecTrim(XRSL("TRIMESTRE"),XrsL("YANIO"),FIRR,FFRR)
        FFINTRR    = FechadorCor(FFRR, "RI")


        IF GENERAIDPRES="S" THEN
           arnom3   = ucase("ID_PRES_" & FFINTRR & ".TXT")
           archivo3 = ruta & "interfases\BCRA\" & arnom3

           if session("APLMDCLI")="BC" THEN

              wwwtdato = "00069000021R"

              strsqlx = "insert into orrepositor (tarchivo,tnombre,tlote,tsecuen,tdatos)"
              strsqlx = strsqlx & " values "
              strsqlx = strsqlx & " ('IDPRESRR','" & arnom3 & "','" & LOTE & "','1','" & wwwtdato & "') "
             'response.write strsqlx

              SET xxxx = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)

           else

              set fname=fs.CreateTextFile(archivo3,true)
              fname.WriteLine("00069000021R")
              fname.Close
              set fname=nothing

           end if

           Sid   = lote

           strsql = "Insert into ORREGINFOBCRAEXT "
           strsql = strsql & " (CODLOTE,ESTADO,TRIMESTRE,ANO,TOTAL,ARCHIVO"
           strsql = strsql & " ) Values ("
           strsql = strsql &  "'" & Sid & "','2','" & WANTTT & "','" & WANTAA & "','1','" & arnom3 & "'"
           strsql = strsql & ")"
        ''  response.write strsql & "<br>"
           call ofv.crearconsultaEx(StrSql,cn,1,volver)


       ELSE
          arnom3=""
       END IF



        arnom1   = ucase("RRIESGOOP_" & FFINTRR & ".TXT")
        archivo1 = ruta & "interfases\BCRA\" & arnom1


        if session("APLMDCLI")="BC" THEN

        else
           set fname=fs.CreateTextFile(archivo1,true)
        end if


        LINEAP     = coddiseno & codentidad & FFINTRR

        X=TRUE

        FHOY       = CLNG(DATE()) + clng(cdate("1-1-1900"))
        Iraiz      = 0
        Isecuencia = 0
        wwwtsecuen = 0

        XTXTCOD  = "RRIESGOOP"
        XARCHCOD = arnom1


    END IF

    CALL GENEVENTOTXT(XRSL,XTXTCOD,XARCHCOD)

    wwTOTRR = wwTOTRR + 1


   Xrsl.movenext
 loop


           if session("APLMDCLI")="BC" THEN

           else
              fname.Close
              set fname=nothing
           ''''   set fs=nothing
           end if

           Sid   = lote

           strsql = "Insert into ORREGINFOBCRAEXT "
           strsql = strsql & " (CODLOTE,ESTADO,TRIMESTRE,ANO,TOTAL,ARCHIVO"
           strsql = strsql & " ) Values ("
           strsql = strsql &  "'" & Sid & "','1','" & WANTTT & "','" & WANTAA & "','" & wwTOTRR & "','" & arnom1 & "'"
           strsql = strsql & ")"
        ''  response.write strsql & "<br>"
           call ofv.crearconsultaEx(StrSql,cn,1,volver)


 end if


   CALL OFV.CERRARCONSULTA(XRSL)









''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''











 IF GENERAIDPRES="S" THEN

    arnom2   = ucase("ID_CD_" & FFINT & ".TXT")
    archivo2 = ruta & "interfases\BCRA\" & arnom2

    if session("APLMDCLI")="BC" THEN


       strsqlx = "insert into orrepositor (tarchivo,tnombre,tlote,tsecuen,tdatos)"
       strsqlx = strsqlx & " values "
       strsqlx = strsqlx & " ('IDCD','" & arnom2 & "','" & LOTE & "','1','" & CODIGOCUP & "') "
      ''response.write strsqlx

       SET xxxx = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)


    else


      set fname=fs.CreateTextFile(archivo2,true)
      fname.WriteLine(CODIGOCUP)
      fname.Close
      set fname=nothing

    end if


 ELSE
   arnom2=""
 END IF







 strsql = "SELECT max(eventoraiz) as eventoraiz,max(secuencia) as secuencia ,"
 strsql = strsql & " SUM(monto) as monto,"
 strsql = strsql & "max(FECinicio) as FECinicio ,"
 strsql = strsql & "max(FECDESCUBRIMIENTO) as FECDESCUBRIMIENTO ,"
 strsql = strsql & "max(FECFINAL) as FECFINAL ,"
 strsql = strsql & "max(FECREGCONTABLE) as FECREGCONTABLE,"
 strsql = strsql & "max(DESCRIPCION) as descripcion,"
 strsql = strsql & "max(TIPOEVENTO) as TIPOEVENTO,"
 strsql = strsql & "max(LINEANEGOCIO) as LINEANEGOCIO,"
 strsql = strsql & "max(SUCURSAL) as SUCURSAL,"
 strsql = strsql & "max(PROCESO) as PROCESO,"
 strsql = strsql & "max(PRODUCTO) as PRODUCTO,"
 strsql = strsql & "max(CARACTER) as CARACTER,"
 strsql = strsql & "max(PARTIDACONTABLE) as PARTIDACONTABLE,"
 strsql = strsql & "max(COBERPOLIZA)as COBERPOLIZA,"
 strsql = strsql & "max(VINCOTROCRIESGO) as VINCOTROCRIESGO ,"
 strsql = strsql & "max(AREASOPORTE) as AREASOPORTE,"
 strsql = strsql & "max(ESTADO) as EVESTADO"
 strsql = strsql & " FROM ORREGINFOBCRADET  WHERE CODLOTE='" &  LOTE & "' AND (ESTADO = 'N' OR ESTADO = 'J') "
 strsql = strsql & " group by eventoraiz,secuencia "
 strsql = strsql & "  order by  eventoraiz,secuencia"
'response.write strsql

 SET XRSL = ofv.crearconsultaEx(StrSql,cn,1,parametros)

 IF GENERAIDPRES="S" THEN
    arnom3   = ucase("ID_PRES_" & FFINT & ".TXT")
    archivo3 = ruta & "interfases\BCRA\" & arnom3

    if session("APLMDCLI")="BC" THEN

       IF xRSL.EOF THEN
          wwwtdato = "00069000010N"
       ELSE
          wwwtdato = "00069000011N"
       END IF


       strsqlx = "insert into orrepositor (tarchivo,tnombre,tlote,tsecuen,tdatos)"
       strsqlx = strsqlx & " values "
       strsqlx = strsqlx & " ('IDPRES','" & arnom3 & "','" & LOTE & "','1','" & wwwtdato & "') "
      'response.write strsqlx

       SET xxxx = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)

    else

       set fname=fs.CreateTextFile(archivo3,true)
       IF XRSL.EOF THEN
          fname.WriteLine("00069000010N")
       ELSE
          fname.WriteLine("00069000011N")
       END IF
       fname.Close
       set fname=nothing

    end if

 ELSE
   arnom3=""
 END IF


 arnom1   = ucase("RIESGOOP_" & FFINT & ".TXT")
 archivo1 = ruta & "interfases\BCRA\" & arnom1


    if session("APLMDCLI")="BC" THEN

    else
       set fname=fs.CreateTextFile(archivo1,true)
    end if


 LINEAP     = coddiseno & codentidad & FFINT

 X=TRUE

 FHOY       = CLNG(DATE()) + clng(cdate("1-1-1900"))
 Iraiz      = 0
 Isecuencia = 0
 wwwtsecuen = 0



 do while not Xrsl.eof



    XTXTCOD  = "RIESGOOP"
    XARCHCOD = arnom1

    CALL GENEVENTOTXT(XRSL,XTXTCOD,XARCHCOD)


   Xrsl.movenext
 loop

    if session("APLMDCLI")="BC" THEN

    else
       fname.Close
       set fname=nothing
       set fs=nothing
    end if

   CALL OFV.CERRARCONSULTA(XRSL)

   strsql="update ORREGINFOBCRA  SET ESTADO=3, archivo1='" & arnom1 & "' ,archivo2='" & arnom2 & "' ,archivo3='" & arnom3 & "' where codlote='" & lote & "'"
 'response.write strsql
   call ofv.crearconsultaEx(StrSql,cn,1,parametros)

end function

FUNCTION GENEVENTOTXT(RSL,TXTCOD,ARCHCOD)

    SECUENCIA = rsL("SECUENCIA")

    IF SECUENCIA = -1 THEN SECUENCIA = 0

    LINEA = LINEAP & strpad(rsL("EVENTORAIZ"),7,"0",2) & strpad(SECUENCIA,3,"0",2)
    LINEA = LINEA  & strpad(rsL("CARACTER"),2,"0",2)

    FECINICIO = rsL("FECINICIO") + clng(cdate("1-1-1900"))

    if clng("0" & rsL("FECINICIO")) = 0 then
       LINEA = LINEA & strpad("0",8,"0",2)
    else
       LINEA = LINEA  & FechadorCor(FECINICIO, "RI")
    end if

    'FECDES= rsL("FECDESCUBRIMIENTO") + clng(cdate("1-1-1900"))
    'FECINICIO= rsL("FECINICIO") + clng(cdate("1-1-1900"))
    'LINEA=LINEA  & FechadorCor(FECINICIO, "RI")

    FECDES = rsL("FECDESCUBRIMIENTO") + clng(cdate("1-1-1900"))

    LINEA = LINEA  & FechadorCor(FECDES, "RI")

    FECREGCONTABLE = rsL("FECREGCONTABLE") + clng(cdate("1-1-1900"))

    if clng("0" & rsL("FECREGCONTABLE")) = 0 then
       LINEA = LINEA & strpad("0",8,"0",2)
    else
       LINEA = LINEA  & FechadorCor(FECREGCONTABLE, "RI")
    end if

    FECFINAL = rsL("FECFINAL") + clng(cdate("1-1-1900"))

    if clng("0" & rsL("FECFINAL"))= 0 then
       LINEA = LINEA & strpad("0",8,"0",2)
    else
       LINEA = LINEA  & FechadorCor(FECFINAL	, "RI")
    end if

    DESCRI = EliminarAcentos(rsL("DESCRIPCION")&"")
    descri = replaceex(descri)
    if descri = "XXX" THEN DESCRI = ""

    LINEA = LINEA  & strpad(DESCRI,300," ",1)
    LINEA = LINEA  & strpad(rsL("TIPOEVENTO"),2,"0",2)
    LINEA = LINEA  & strpad(rsL("LINEANEGOCIO"),2,"0",2)

    SUCURSAL = EliminarAcentos(rsL("SUCURSAL") & "")
    sucursal = replaceex(SUCURSAL)
    LINEA = LINEA  & strpad(SUCURSAL,50," ",1)

    PROCESO = EliminarAcentos(rsL("PROCESO") & "")
    PROCESO = replaceex(PROCESO)
    LINEA = LINEA  & strpad(PROCESO,50," ",1)

    PRODUCTO = EliminarAcentos(rsL("PRODUCTO") & "")
    PRODUCTO = replaceex(PRODUCTO)
    LINEA = LINEA & strpad(PRODUCTO,50," ",1)

    MONTO = Clng(round(RSL("MONTO"),0))

    IF MONTO < 0 THEN
       MONTO =  MONTO * -1
       IF RSL("CARACTER") = "06" OR RSL("CARACTER") = "07" THEN
          LINEA = LINEA  & strpad(MONTO,10,"0",2)
       ELSE
          LINEA = LINEA  & "-" & strpad(MONTO,9,"0",2)
       END IF

    ELSE

       IF RSL("CARACTER") = "06" OR RSL("CARACTER") = "07" THEN
          LINEA = LINEA  & "-" & strpad(MONTO,9,"0",2)
       ELSE
          LINEA = LINEA  & strpad(MONTO,10,"0",2)
       END IF

    END IF

    LINEA = LINEA & strpad(RSL("PARTIDACONTABLE"),6,"0",2)

    COBERPOLIZA = RSL("COBERPOLIZA")

    PAIN  = CLNG(COBERPOLIZA)
    PADEC = PAIN - COBERPOLIZA

    LINEA = LINEA & strpad(PAIN,3,"0",2)  & strpad(PADEC,2,"0",2)
    LINEA = LINEA & strpad(RSL("VINCOTROCRIESGO"),1,"0",2)
    LINEA = LINEA & strpad(RSL("AREASOPORTE"),1,"0",2)

    if session("APLMDCLI") = "BC" THEN

       wwwtsecuen = wwwtsecuen + 1

       wwwtdato = LINEA

       strsqlx = "insert into orrepositor (tarchivo,tnombre,tlote,tsecuen,tdatos)"
       strsqlx = strsqlx & " values "
       strsqlx = strsqlx & " ('" & TXTCOD & "','" & ARCHCOD & "','" & LOTE & "','" & wwwtsecuen & "','" & wwwtdato & "') "
      'response.write strsqlx

       SET xxxx = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)


    else
       fname.WriteLine(LINEA)
    end if



END FUNCTION



function replaceex(textoz)
 FOR CCC=1 TO LEN(textoz)

  WKFORDESCRICCC = MID(textoz,CCC,1)
  '  IF WKFORDESCRICCC > CHR(125) or WKFORDESCRICCC < CHR(32) THEN
     if (WKFORDESCRICCC >= chr(48) and WKFORDESCRICCC <= chr(57)) or (WKFORDESCRICCC >= chr(65) and WKFORDESCRICCC <= chr(90)) then
     else
      WKFORDESCRICCC= " "
    end if
  replaceex=replaceex & WKFORDESCRICCC
 next


end function

Response.Write sHTML
%>