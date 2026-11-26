
<%

 FUNCTION GENFMTOBJDSC(FNCODOBJ,FNTIPOBJ)


  FNSHTML = ""


  FNSHTML = FNSHTML & ofv.GenTabla("","","","100% valign=middle margin=3  ","",0,0,0,0,"","","","no")




  stitulo = " "
  smuestra = "NO"
  sdescri = " "


  Strsql = "Select muestra,titulo,descri From Objdscabr Where codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rs.eof then
     stitulo = rs(1)
     smuestra = rs(0)
     sdescri = rs(2)
  end if
  call ofv.cerrarconsulta(rs)
  'call ofv.cerrarconn(cn)

'  FNSHTML = FNSHTML & ofv.GenTabla("","","","80%","","0","0","1","0","","","","no")

  if isnull(smuestra) then smuestra = "NO"
  if isnull(sdescri) then sdescri = " "

'Objetivo
  if smuestra = "NO" then

     h1 = ofv.GenCelda("","ttt","","","","","","Sin&nbsp;Descripcion","si")
    FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")

  else


     h1 = ofv.GenCelda("","tt style={border-style:outset;border-width:2;} ","","","","","","<font face=tahoma >" & stitulo & ":</font>","si")
    FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")
     h1 = ofv.GenCelda("","cc style=padding:5px;","","","","","","<p align=justify  style='{font-family:tahoma;font-size:10pt;}' >" & replace(sdescri,chr(13),"</p><p align=justify style='{font-family:tahoma;font-size:10pt;}' >") & "</p>","si")
    FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")


  end if


  FNSHTML = FNSHTML & ofv.GenFTabla()


  GENFMTOBJDSC = "<div id=zglos class='norma-card'>" & FNSHTML  & "</div>"


 END FUNCTION


%>