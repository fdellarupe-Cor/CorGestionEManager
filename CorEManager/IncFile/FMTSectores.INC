
<%

 FUNCTION GENFMTSECTORES(FNCODOBJ,FNTIPOBJ)


  FNSHTML = ""


  FNSHTML = FNSHTML & ""

  'sImagen4si = "<FONT FACE=corflaticon COLOR=#f7aa22 SIZE=3>G</FONT>"
  sImagen4si = "<i class='las la-university' ></i>"

  'Encabezado sector
  FNSHTML = FNSHTML & "<table width=100% cellspacing=0 cellpadding=3   >"
  FNSHTML = FNSHTML & "<tr height=30 ><td width=100% "
  FNSHTML = FNSHTML & "class=ttt align=center colspan=3 >SECTORES AFECTADOS</td></tr>"


  Strsql = "SELECT ECO.descri, OBJSECTOR.*  FROM OBJSECTOR, ECO "
  Strsql = Strsql & "where OBJSECTOR.codsector = ECO.eco and OBJSECTOR.codobjeto=" & FNCODOBJ
  Strsql = Strsql & " Order By Calificador Desc"
 'response.write strsql
 if Cobjun=1 and Session("EMPMUN") = "S" then
    	  sql="exec SpSectoresCorp " & FNCODOBJ &",'" & FNTIPOBJ & "'," & session("un")

    	  Set rsSecCorp = ofv.crearconsultaEx(sql,cn,1,parametros)

            if not rsSecCorp.eof then

                ASecCorp =rsSecCorp.getrows
             end if
           Call ofv.cerrarconsulta(rsSecCorp)
     end if

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs.eof



    ' estado observacion

	      Celdaobs  = ""


	      if Nestado=5  then




	               sAccion = "genObsdoc.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
	      	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & rs(1)   & "&ID2="& rs(1) &"&ID3=" & rs(1)

	      	       xsffun  = " onclick=""document.getElementById('DX4').style.display='';"" "
	      	       sHref   = sAccion
	      	       sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
''	      	       Celdaobs  = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","","","1","","",sInput,"si")
''	               Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",Celdaobs,"si")

Celdaobs= ofv.GenLink("","ut" & "  ",sHref,"WK4","","",sInput,"si")
Celdaobs = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","","","1","","",Celdaobs,"si")




	      end if




    FNSHTML = FNSHTML & "<tr height=35 >"

    FNSHTML = FNSHTML & "<td  class=uts  >"
    FNSHTML = FNSHTML & sImagen4si & ""
    FNSHTML = FNSHTML & "</td>"

      'Texto sectores origen

    FNSHTML = FNSHTML & Celdaobs & "<td  class=ut width=98% >" '& Cobjun
    if Cobjun=1 and Session("EMPMUN") = "S" then
     if isarray(ASecCorp) and Cobjun=1 and Session("EMPMUN") = "S" then
	   SECD  = rs(0)
	   SECD = FindInArray(ASecCorp,6,rs("codsector"),1,SECD)
	   SArma = "" & SECD & ""
	 END IF
    else
      SArma = "" & rs(0) & ""

    end if
    FNSHTML = FNSHTML & SArma

    FNSHTML = FNSHTML & "</td></tr>"

    rs.movenext
  loop

  'FNSHTML = FNSHTML & "<tr ><td  class=vt >&nbsp;</td></tr>"
  FNSHTML = FNSHTML & "</table>"

  call ofv.cerrarconsulta(rs)

  FNSHTML = FNSHTML & ""

  GENFMTSECTORES = "<div id=zglos class='norma-card'>" & FNSHTML  & "</div>"


 END FUNCTION



%>