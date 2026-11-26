<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
 dim arrfontx(10,1)

  response.expires=0

  CantRegMostrarx = 10
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

 XXXX = NEWEMPRE("NO")

      session("ASP")="SUBCONSULTA.ASP"

      sUN = request.querystring("EMPRE")
    'response.write "sEMPRE: " & sEMPRE
       IF sUN = ""  AND SESSION("UN") = "" THEN

         sUN = ATREMP(0, 0)
         SESSION("UN") = sUN & ""
         If (SESSION("UN") & "") = "" Then SESSION("UN") = 1
       elseIF sUN <> "" THEN
        XXXX = NEWEMPRE(sUN)
        SESSION("UN")=sUN
       ELSE
        XXXX = NEWEMPRE(SESSION("UN"))
        sUN=SESSION("UN")
   end if

'response.write session("UN")

  sHTML = ofv.FormHeader(session("FormN"))

 TablaH = ofv.GenTabla("","","","20%","","0","0","0","25","","",selecEmpre(),"si")


shtml = shtml & tablah

   'Encabezados
  Formulario = ofv.Encabezados("id","Titulo","activo","cant&nbsp;reg","ORDEN","WIDTH","HEIGHT")',"ICONO","MOSTRAR"
 ' Formulario = ofv.Encabezados("id","Titulo","activo","cant&nbsp;reg")',"ICONO","MOSTRAR"




  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From Novedtitex where codemp=" & session("UN") & " order by orden"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
  posi=0




'    Celda = ofv.GenCelda("","tt","","","","","","Tipo:","si")
'    Celda = Celda & ofv.GenCelda("","tt","center","","","","","Titulo:","si")
'    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

   'for tt=2 to 10

    	arrfontx(0,0)="lar la-star"
   	   	arrfontx(0,1)="star"
   	   	arrfontx(1,0)="lar la-file-alt"
	   	arrfontx(1,1)="file"
	   	arrfontx(2,0)="lar la-comment"
	   	arrfontx(2,1)="comment"
	   	arrfontx(3,0)="las la-hourglass-half"
	   	arrfontx(3,1)="hourglass"
	   	arrfontx(4,0)="las la-bullhorn"
	   	arrfontx(4,1)="bullhorn"
	   	arrfontx(5,0)="lar la-clock"
	   	arrfontx(5,1)="clock"
	   	arrfontx(6,0)="las la-book"
	   	arrfontx(6,1)="book"

 do while not rs.eof
           sForm = "form" & posi
           fila  = ""
          'Celda = ofv.GenCelda("","tt","","","","","","&nbsp;","si")
         Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")

          Celda = celda & ofv.GenCelda("","cc","rigth","","","","",rs(5),"si")

          sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"120","dt","")
          Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")


           scomK = "S" & chr(9) & "N"
		   scomD = "SI" & chr(9) & "NO"


	  if bModificar then
		  sInput = ofv.gencomboFX(sForm,rs(2).name,rs(2),scomK,scomD," ")
	  else
		  sInput = ofv.GenerarInput("",rs(2),"readonly","","10","dt","")
	  end if
      Celda =  Celda & ofv.gencelda("","vti","","","","","",sInput,"si")



       sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"20","dtn","")
	   Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

	   sInput = ofv.GenerarInput(rs(6).name,rs(6),"text",sForm,"20","dtn","")
	   Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
	   sInput = ofv.GenerarInput(rs(7).name,rs(7),"text",sForm,"20","dtn","")
	   Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

	   sInput = ofv.GenerarInput(rs(8).name,rs(8),"text",sForm,"20","dtn","")
	   Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
	   'sxopc1 =  "lar la-star" &CHR(34)&




	   sxopc1=""
	   MOSTRAICONO=""
	   for i = 0 to 5

	   if UCASE(rs("icono") &"")= UCASE(arrfontx(i,0) &"") then
	     sxopc1 =  sxopc1 &  "<option selected value='" & arrfontx(i,0) & "' > " & arrfontx(i,1)&" </option> "

	   else
	    sxopc1 =  sxopc1 &  "<option value='" & arrfontx(i,0) & "' > " & arrfontx(i,1)&" </option> "

	   end if
	   next
	'   sxopc1 =  "<option value='las la-hourglass-half' > hourglass </option> "
	'   sxopc1 =  sxopc1 &  "<option value='lar la-file-alt' > file </option> "
	   'Chr(9) & "<i class=""lar la-file-alt""></i>" & Chr(9) & "<i&nbsp;class=""lar&nbsp;la-comment""></i>"
	'   sCOMK = sCOMK  & Chr(9) & "<i class='las la-hourglass-half'></i>" & Chr(9) & "<i&nbsp;class='las&nbsp;la-bullhorn'></i>" & Chr(9) & "<i&nbsp;class='lar&nbsp;la-clock'></i>"
	'   sCOMK = sCOMK  & Chr(9) & "< class='las la-book'></i>"
	 '  sCOMD = "star" & Chr(9) & "file" & Chr(9) & "comment"& Chr(9) & "hourglass"& Chr(9) & "bullhorn"& Chr(9) & "clock"& Chr(9) & "book"

	   sInput = "<select onchange='document."& sForm &".submit();' class=at id='" & rs(11).name & "' name='" & rs(11).name & "' >" & sxopc1 & "</select>"
       'Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

       sInput = "<i class='" & RS("ICONO")& " la-2x'></i>"
      ' Celda = Celda & ofv.GenCelda("","","CENTER","","","","",sInput,"si")
        Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")



    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=NovedTitex&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")
    rs.movenext
    posi=posi+1
 loop
  end if

  ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa ","","","","0","","0","0","","",filas,"si")


  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no","",Buscar,celda)



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if


function selecEmpre()

sImagenx = "<B><FONT FACE=CorFlatIcon COLOR=#aa0000 SIZE=4>x</FONT></B>"

   IF  Session("EMPMUN") = "S" THEN

 '  Fila    = ""

  ' SINPUT3 = "&nbsp;<A class=tt style='color:#000000;' title='SELECCIONE EMPRESA' HREF='javascript:void(0);' onclick='shwDem();' >EMPRESA:&nbsp;&nbsp;</a>"

  ' Fila    = ""
  ' Celda   = ""
 '  sHref   = ""

  ' sImage  = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
  ' sLink   = ofv.GenLink("","ttt",sHref,"ftemas","","",sImage,"si")
   'Celda   = Celda & ofv.GenCelda("","tt ","","8%","","","",SINPUT3,"si")
   'Celda   = Celda & ofv.GenCelda("","ttt ","center","5%","","","",sLink,"si")
   'Celda   = Celda & ofv.GenCelda("","ttt  ","","","","","","&nbsp;<BIG>" & ATREMP(0,4) & "</BIG>","si")

  ' Fila    = fila & ofv.GenRow("","","top","","","","",Celda,"si")

   'Tabla   = ofv.GenTabla("","aaa style='"  & Session("CBRDABL") & "' ","","100% ","","0","","-1","-2","","",Fila,"si")





   'shtml=shtml & "<script language=javascript>" & chr(13)
  ' shtml=shtml & "function shwDem(){" & chr(13)
  ' shtml=shtml & "document.all.DEM1.style.display='block';" & chr(13)
   'shtml=shtml & "}" & chr(13)
   'shtml=shtml & "</script>" & chr(13)
 '----------------------------------------------
 '  sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
  ' sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
  ' sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "document.all.DEM1.style.display='none';" & chr(34)
  ' sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
  ' TABLA = Tabla & "<DIV id='DEM1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:240px;left:20px;height:250px;width:350px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"


 '  TABLAX = TablaX & OFV.gentabla("", "aaa style='filter:shadow(color=#dddddd,direction=125);' ", "", "100%", "", "0", "", "-2", "-1", "", "", "", "no")
 '  TABLAX = TablaX & "<tr><td  "
 '  TABLAX = TablaX & "class=vT align=left colspan=3 STYLE='border-style:outset;border-width:2;background-color:#eeeeee;' >" & sinput2 & "&nbsp;SELECCIONE EMPRESA</td></tr>"
 ''  TABLAX = TablaX & "<tr><td colspan=3 "
 '  TABLAX = TablaX & " >&nbsp;</td></tr>"

'   FOR IAEMP=1 TO TAEMP

    '''  RESPONSE.WRITE  ATREMP(IAEMP,2)

 '     IF ATREMP(IAEMP,2) = 1 THEN

 '      IF ATREMP(IAEMP,0) = ATREMP(0,0) THEN
 '         TABLAX = TablaX & "<tr BGCOLOR=#EEEEEE >"
 '      ELSE
  '        TABLAX = TablaX & "<tr>"
  '     END IF

 '      TABLAX = TablaX & "<td width=5 "
  '     IF ATREMP(IAEMP,0) = ATREMP(0,0) THEN
 '         TABLAX = TablaX & "class=vt  >" & sImagen4si & "</td>"
  '     ELSE
  '        TABLAX = TablaX & "class=vt  >" & sImagenI3 & "</td>"
  '     END IF

'       sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(IAEMP,5),"0","")
  '     TABLAX = TablaX & "<td width=5 "
'       TABLAX = TablaX & "class=vt  >" & sImage & "</td>"
'
  '     TABLAX = TablaX & "<td class=vt  >"

       'sHref = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0&EMPRE=" & ATREMP(IAEMP,0)
 ''		sHref = "NovedTit.asp?EMPRE=" & ATREMP(IAEMP,0)
     '''  sHref = "OBJtreeTema.asp?Tipo=" & sTipo & "&Nombre=" & LNombre & "&Menu=" & SESSION("MENU") & "&EMPRE=" & ATREMP(IAEMP,0)
  '     sLink = ofv.GenLink("","cc",sHref,"","","","<b>" & ATREMP(IAEMP,4) & "</b>","si")
   ''    TABLAX = TablaX & sLink
   '    TABLAX = TablaX & "</td>"
    '   TABLAX = TablaX & "</tr>"
   '   END IF

  ' NEXT

  ' TABLAX = TablaX & OFV.genFtabla()
'
  ' TABLA = Tabla & TABLAX

   'TABLA = Tabla & "</DIV>"
    TableEmp=GenEmpresas(ATREMP,"NOVEDTIT","")
	sHTML = sHTML & TableEmp
   'sHTML = sHTML & TABLA

  end if

end function

Response.Write sHTML

%>