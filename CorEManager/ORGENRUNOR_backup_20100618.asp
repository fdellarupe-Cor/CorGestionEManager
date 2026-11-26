<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  xxx=APLCHECK()

  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  sopcion  = request.querystring("opcion") & ""

  sHTML = ofv.FormHeader(session("FormN"))



    sInput = "ADMINISTRACION DE LOTES DE EVENTOS DE PERDIDA A BCRA"
    Celda = ofv.gencelda("","TT ","","90%","","","",sInput,"si")


    Fila = ofv.genrow("","","","","","","",Celda,"si")

  Formulario = Fila
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si") & "<BR>"

  shtmlz=sHTML

  Formulario = ""





  'Encabezados
  Formulario = ofv.Encabezados("LOTE","ESTADO","DESCRIPCION","TRIMESTRE","AÑO","ULT.EVENTO","&nbsp;","&nbsp;","&nbsp;","&nbsp;","&nbsp;","&nbsp;")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


  wwkano = 2010
  wwktri = 1


'''''''''''''''''''''''''''''''''''''''
'' INCORPORO LOTE NUEVO

  if sopcion = "SI" then

     sfechai = clng(date()) - clng(cdate("01-01-1900"))

     Strsql = "Select max(ano) From ORREGINFOBCRA "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then 
        wwkano = rs(0)

        Strsql = "Select max(trimestre) From ORREGINFOBCRA where ano = " & wwkano
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rs.eof then 
           wwktri = rs(0)

           if wwktri < 4 then
              wwktri = wwktri + 1
           else
              wwktri = 1
              wwkano = wwkano + 1
           end if


        end if


     end if


     STRSQL="SELECT * FROM ORCONFIGBCRA"
     SET RS=ofv.crearconsultaEx(StrSql,cn,1,volver)


     IF not rs.eof then
       Sid= "RI" & GTKEY("")
        iTRIM=GetTrimestre(DATE())
        strsql = "Insert into ORREGINFOBCRA "
        strsql = strsql & " (CODLOTE,CODDISENO,CODENTIDAD,ESTADO,TRIMESTRE,ANO,ULTEVENTO"
        strsql = strsql & " ) Values ("
        strsql = strsql &  "'" & Sid & "','" &  rs("coddiseno") & "','"&rs("codentidad")&"','0','" & wwktri & "'"
        strsql = strsql & "," & wwkano & ",'0')"
      '  response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)
     end if



     call ofv.cerrarconsulta(rs)
  end if



  STRSQL="SELECT * FROM ORCONFIGBCRA"
  SET RS=ofv.crearconsultaEx(StrSql,cn,1,volver)

  if len(trim(rs("codentidad")))>0 then
   xva=true
  else
   xva=false
  end if

'  response.write len(trim(rs("codentidad")))
  call ofv.cerrarconsulta(rs)

  WKSETTABLA = "ORREGINFOBCRA"


  

  Strsql = "Select * From " & WKSETTABLA & " order by ano desc, trimestre desc, estado , objid desc"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0
  wwkagr = "S"
  Do While (Not rs.EOF)  and xva
    sForm = "form" & posi




    Formu1 = "" 
    Formu2 = "" 
    Formu3 = "" 
    Formu4 = "" 
    Formu5 = "" 

    wkestado = rs("ESTADO") & ""

    'Eliminar
    if bSuprimir AND wkestado < 2 then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(0),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'CODIGO
    sInput = RS(0) & "&nbsp;"
    Celda = Celda & ofv.gencelda("","vti STYLE='border-style:solid;border-width:1;border-color:#000000;' BGCOLOR='#DDDDFF' " ,"","5%","","","",sInput,"si")



   'ESTADO

    if wkestado < "4" then
       wwkagr = "N"
    end if


    select case wkestado

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
    Celda = Celda & ofv.gencelda("","vT STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","10%","","","",sInput,"si")
   IF wkestado < 4 and posI = 0 THEN
    sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"40","dt","")
    CELDAX=ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;' " ,"","25%","","","",sInput,"si")
   ELSE
    CELDAX = ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","25%","","","",RS(2)&"&nbsp;","si")
   END IF

   Celda = Celda & CELDAX
    iTRIMESTRE=RS("TRIMESTRE")
    SCOMSEP= CHR(9)
  	 scomK=""
  	 scomD=""
  	 FOR I =1 TO 4
  	   IF RS("ANO") = YEAR(DATE()) or session("APLMDCLI")="BP" THEN
  	    scomK   = scomK & I   & SCOMSEP
  	    scomD   = scomD & "TRIMESTRE " & I  & SCOMSEP
  	   ELSE
  	    scomK   = scomK & 4   & SCOMSEP
	    scomD   = scomD & "TRIMESTRE " & 4  & SCOMSEP
	     STRSQL="UPDATE ORREGINFOBCRA SET TRIMESTRE=4 WHERE CODLOTE='" & RS("CODLOTE") &"'"
         iTRIMESTRE=4
	     CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
  	   EXIT FOR
  	  END IF
  	 NEXT

     ' scomD=REPLACE(scomD," ","&nbsp;")

   IF wkestado < 0 and posI = 0 THEN

    sInput = ofv.gencomboFX(sForm,"TRIMESTRE",ITRIMESTRE,scomK,scomD,"si")
   ELSE
    SINPUT ="TRIMESTRE&nbsp;" & ITRIMESTRE
   END IF

    Celda = Celda & ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","5%","","","",SINPUT,"si")

      ano1=year(date())
      ano2=ano1-1
      SCOMSEP= CHR(9)
      scomK=""
      scomD=""
	  scomK   = scomK  & ano1  & SCOMSEP & ano2
	  scomD   = scomD &  ano1   & SCOMSEP & ano2


	     ' scomD=REPLACE(scomD," ","&nbsp;")

	 IF wkestado < 0  and posI = 0  THEN
	    sInput = ofv.gencomboFX(sForm,"ano",RS("ano"),scomK,scomD,"si")
	 ELSE
	    SINPUT= RS("ano")
	 END IF

    Celda = Celda & ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","5%","","","",SINPUT,"si")

    SINPUT= RS(23) & "&nbsp;"
    Celda = Celda & ofv.gencelda("","VT  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","right","5%","","","",SINPUT,"si")




if wkestado < 4 and posI = 0 then
  sAccion = "ORGENRUNOREX.asp?LOTE=" & RS("CODLOTE")
  BTIT1 = "<INPUT TYPE='BUTTON' VALUE='PROCESAR' CLASS='cc' title='Procesar el lote.' STYLE='cursor:pointer;background-color:#008800;color:#ffffff;' onclick=""disp_confirm('Confirma el procesamiento de este lote???','FINI" & posi & "');""  > "
  Formu1 = ofv.GenForm("FINI" & posi,"",sAccion,"","post","","","si")
  Celda = Celda & ofv.gencelda("","VT title='PROCESAR LOTE.' STYLE='color:#ffffff;border-style:solid;border-width:1;border-color:#888888;'","CENTER","5%","","","",BTIT1,"si")


elseif wkestado = 4 and posI = 0 then
  sAccion = "ORGENRUNORrevers.asp?LOTE=" & RS("CODLOTE")
  BTIT1 = "<INPUT TYPE='BUTTON' VALUE='REVERTIR' CLASS='cc' title='Revertir el lote.' STYLE='cursor:pointer;background-color:#990000;color:#ffffff;' onclick=""disp_confirm('Confirma la reversion de este lote???','FINI" & posi & "');""  > "
  Formu1 = ofv.GenForm("FINI" & posi,"",sAccion,"","post","","","si")
  Celda = Celda & ofv.gencelda("","VT title='REVERTIR LOTE.' STYLE='color:#ffffff;border-style:solid;border-width:1;border-color:#888888;'","CENTER","5%","","","",BTIT1,"si")



else


  Celda = Celda & ofv.gencelda("","ut STYLE='border-style:solid;border-width:1;border-color:#888888;'","CENTER","5%","","","","-----","si")


end if





 if wkestado > 1  AND wkestado < 5 then
  sAccion = "ORRUNORLOTEDET.asp?LOTE=" & RS("CODLOTE")
  BTIT1 = "<INPUT TYPE='BUTTON' VALUE='SELECCIONADOS' CLASS='cc' title='Detalle del lote.' STYLE='cursor:pointer;background-color:#000099;color:#ffffff;' onclick=""disp_confirm('Confirma Ver el detalle de este lote???','FOLOTEDET" & posi & "');""  > "
  Formu4 = ofv.GenForm("FOLOTEDET" & posi,"",sAccion,"","post target=_new","","","si")
  Celda = Celda & ofv.gencelda("","VT title='DETALLE DEL LOTE.' STYLE='color:#ffffff;border-style:solid;border-width:1;border-color:#888888;'","CENTER","5%","","","",BTIT1,"si")

 else

  Celda = Celda & ofv.gencelda("","ut STYLE='border-style:solid;border-width:1;border-color:#888888;'","CENTER","5%","","","","-----","si")

 end if



 if wkestado > 1  AND wkestado < 5 then
  sAccion = "ORRUNORLOTEERR.asp?LOTE=" & RS("CODLOTE")
  BTIT1 = "<INPUT TYPE='BUTTON' VALUE='NO SELECCIONADOS' CLASS='cc' title='ERRORES DEL LOTE.' STYLE='cursor:pointer;background-color:#000099;color:#ffffff;' onclick=""disp_confirm('Confirma ver errores de este lote???','FINIERR" & posi & "');""  > "
  Formu2 = ofv.GenForm("FINIERR" & posi,"",sAccion,"","post target=_new","","","si")
  Celda = Celda & ofv.gencelda("","VT title='ERRORES DEL LOTE.' STYLE='color:#ffffff;border-style:solid;border-width:1;border-color:#888888;'","CENTER","5%","","","",BTIT1,"si")

 else

  Celda = Celda & ofv.gencelda("","ut STYLE='border-style:solid;border-width:1;border-color:#888888;'","CENTER","5%","","","","-----","si")

 end if



 if wkestado > 1  AND wkestado < 5 then
   sAccion = "ORRUNORLOTEINF.asp?LOTE=" & RS("CODLOTE")
   BTIT1 = "<INPUT TYPE='BUTTON' VALUE='TOTALES' CLASS='cc' title='TOTALES DEL LOTE.' STYLE='cursor:pointer;background-color:#000099;color:#ffffff;' onclick=""disp_confirm('Confirma ver TOTALES de este lote???','FINIX" & posi & "');""  > "
   Formu3 = ofv.GenForm("FINIX" & posi,"",sAccion,"","post target=_new","","","si")
   Celda = Celda & ofv.gencelda("","VT title='TOTALES DEL LOTE.' STYLE='color:#ffffff;border-style:solid;border-width:1;border-color:#888888;'","CENTER","5%","","","",BTIT1,"si")

 else

  Celda = Celda & ofv.gencelda("","ut STYLE='border-style:solid;border-width:1;border-color:#888888;'","CENTER","5%","","","","-----","si")

 end if


 if wkestado = 3 then
   sAccion = "ORGENRUNORrevers.asp?LOTE=" & RS("CODLOTE") & "&OPCION=FIN" 
   BTIT1 = "<INPUT TYPE='BUTTON' VALUE='FINALIZAR' CLASS='cc' title='Marcar lote como enviado y finalizar.' STYLE='cursor:pointer;background-color:#990000;color:#ffffff;' onclick=""disp_confirm('Confirma marcar lote como enviado y finalizar???','FINIFFF" & posi & "');""  > "
   Formu5 = ofv.GenForm("FINIFFF" & posi,"",sAccion,"","post","","","si")
   Celda = Celda & ofv.gencelda("","VT title='Marcar lote como enviado y finalizar.' STYLE='color:#ffffff;border-style:solid;border-width:1;border-color:#888888;'","CENTER","5%","","","",BTIT1,"si")

 else

   Celda = Celda & ofv.gencelda("","ut STYLE='border-style:solid;border-width:1;border-color:#888888;'","CENTER","5%","","","","-----","si")

 end if

 if wkestado > 1  AND wkestado < 1 then
  sAccion = "ORGENRUNOREXXX.asp?LOTE=" & RS("CODLOTE")
  BTIT1 = "<INPUT TYPE='BUTTON' VALUE='CHECK LOTE " & RS(0) & "' CLASS='cc' title='check lote.' STYLE='cursor:pointer;background-color:#008800;color:#ffffff;' onclick=""disp_confirm('Confirma el procesamiento de este lote???','FINIXXX" & posi & "');""  > "
  Formu9 = ofv.GenForm("FINIXXX" & posi,"",sAccion,"","post target=_new","","","si")
  Celda = Celda & ofv.gencelda("","VT title='CHECK LOTE.' STYLE='color:#ffffff;border-style:solid;border-width:1;border-color:#888888;'","CENTER","5%","","","",BTIT1,"si")
 else

  Celda = Celda & ofv.gencelda("","ut STYLE='border-style:solid;border-width:1;border-color:#888888;'","CENTER","5%","","","","&nbsp;","si")

 end if


 Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=" & WKSETTABLA & "&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si") & Formu1 & Formu2 & FORMU4  & Formu3  & Formu5 & Formu9

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar and   wwkagr = "S" then
    sAccion = "ORGENRUNOR.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri From " & WKSETTABLA  & " order by descri"  ' orden pagina
  str2 = "Select descri From " & WKSETTABLA   & " order by descri"    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))


  Celda = ""

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

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

   if xva then
      Response.Write sHTML
   else
     Celda = ""

     Celda = Celda & ofv.GenCelda("","Vt BGCOLOR='#F7EEDD' ","center","","","","","DEBE INGRESAR CODIGO DE ENTIDAD EN PARAMETROS","si")
     Fila  = ofv.GenRow("","","","","","","",Celda,"si")
     sHTML =  shtmlz &  "<br><br>" & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")
     Response.Write shtml
end if



%>