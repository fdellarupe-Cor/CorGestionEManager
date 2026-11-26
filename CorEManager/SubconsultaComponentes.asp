<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  sMenu   = request.querystring("Menu")
  sNombre = request.querystring("Nombre")
  sTema   = request.form("codtema")
  if sTema = "" or sTema = "nsnc" then sTema = 0
  sSubtema = request.form("codprodserv")
  if sSubtema = "" or sSubtema = "nsnc" then sSubtema = 0
  sEsquema = request.form("codobjeto")
  if sEsquema = "" or sEsquema = "nsnc" then sEsquema = 0
  xComp    = ""
  xComp    = request.form("XComplem")

  'response.write request.form

  set cn = ofv.conectar(ofv.strconn2)

  select case UCase(sNombre)
         case "TIPO"
               Strsql = "Select * From Tipocomp ORDER BY DESCRI "
         case "SISTEMA"
               Strsql = "select DISTINCT sistorigen from Compprty order by sistorigen"
         case "PORTEXTO"

         case "PORIDENTIFICACION"

  end select

  sNombreN = sNombre
  sNombre  = "Consulta por " & sNombre
  snombre  = ucase(snombre)

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

  sHTML = ofv.FormHeader(session("FormN"))

TablaH = ofv.GenTabla("","","","20%","","0","0","0","25","","",selecEmpre(),"si")


	shtml = shtml & tablah

    'sInput = "<big>SELECCIONE LOS PARAMETROS PARA LA CONSULTA</big>"
    'Celda  = ofv.GenCelda("","ttt colspan=2 ","","","","","",sInput,"si")
    'Fila   = ofv.genrow("","","","","30","","",Celda,"si")

    'Celda  = ofv.GenCelda("","ut","","30%","","","","&nbsp;","si")
    'Celda  = Celda & ofv.GenCelda("","uti","","70%","","","","&nbsp;","si")
    'Fila   = Fila & ofv.GenRow("","","","","10","","",Celda,"si")

    'Celda  = ""

'response.write request.querystring & "<br>" & request.form & "<br>" & "NOM: " & sNombre

  if sNombre = "CONSULTA POR TIPO" then



    Celda  = Celda & ofv.GenCelda("","tc","","","","","","Tipo&nbsp;&nbsp;","si")
    'sInput = ofv.GenerarCombox2(strsql,cn,0,1,"Lista") 'ofv.GenerarLista(Strsql,cn,1)
     'sInput = ofv.GenerarCombox2(strsql,cn,0,1,"Lista") 'ofv.GenerarLista(Strsql,cn,1)
    'sInput = ofv.GenerarCombox2(Strsql,cn,0,1,"tipoaccion")
       Set rsPerfil = ofv.crearconsultaEx(Strsql, cn, 1, parametros)

       scomK = ""'"null" & Chr(9)


       scomD = "" '"" & Chr(9)
       SCOMSEP =""
          tttt=""
        if  not rsperfil.eof then
          tttt=rsperfil(0)
        end if
       'if not rsperfil.eof then wwww = rsperfil(0)
       do while not rsperfil.eof
       scomK=scomK & SCOMSEP & rsperfil(0)
       scomd=scomd & SCOMSEP & rsperfil(1)
       SCOMSEP = chr(9)
       rsPerfil.movenext
        loop
                             'id del form, nombre del dato, el dato,'
      sInput =  ofv.gencomboFX("Subconsulta", "Lista",tttt, scomK, scomD, "no0")
    Celda  = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
    Fila   = Fila & ofv.GenRow("","","","","","","",Celda,"si")


  elseif sNombre = "CONSULTA POR IDENTIFICACION" then




    Celda  = Celda & ofv.GenCelda("","tc","","","","","","Identificacion&nbsp;&nbsp;","si")
    sInput = ofv.GenerarInput("buscar","","agregar","","60","dt","")
    Celda  = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
    Fila   = Fila & ofv.GenRow("","","","","","","",Celda,"si")


  elseif sNombre = "CONSULTA POR TEXTO" then


    Celda  = Celda & ofv.GenCelda("","tc","","","","","","Texto&nbsp;&nbsp;","si")
    sInput = ofv.GenerarInput("buscar","","agregar","","60","dt","")
    Celda  = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
    Fila   = Fila & ofv.GenRow("","","","","","","",Celda,"si")

  elseif sNombre = "CONSULTA POR SISTEMA" then



        Celda = Celda & ofv.GenCelda("","tc","","","","","","Sistema&nbsp;&nbsp;","si")
        'sInput = ofv.GenerarLista(Strsql,cn,0)



         Set rsSectorn = ofv.crearconsultaEx(strsql, cn, 1, parametros)
            scomK = ""'"null" & Chr(9)
        scomD = "" '"" & Chr(9)
        SCOMSEP =""
        ht=""
      do while not rsSectorn.eof
         scomK=scomK & SCOMSEP & rsSectorn(0)
         if ht = "" then ht=rsSectorn(0)
         scomd=scomd & SCOMSEP & rsSectorn(0)
         SCOMSEP = chr(9)
         rsSectorn.movenext
        loop
        sInput = ofv.gencomboFX("","lista",ht,scomK,scomD,"no")   'EDITAR! SE TOMA OTRO VALOR si es una VARIABLE
        Celda  = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
        Fila   = Fila & ofv.GenRow("","","","","","","",Celda,"si")

  end if

    'Celda  = ofv.GenCelda("","ut colspan=2","","","","","","&nbsp;","si")
    'Fila   = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    Celda  = ofv.GenCelda("","ut colspan=2","","","","","","&nbsp;","si")
    Fila   = Fila & ofv.GenRow("","","","","","","",Celda,"si")



    sInput = ofv.GenerarInput("","Aceptar","submit","","10","bt","")
    Celda  = ofv.GenCelda("","ut","Center colspan=2","","","","",sInput,"si")
    Fila   = Fila & ofv.GenRow("","","","","","","",Celda,"si")
    sAccion = "ListadoConsultasComp.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0" & "&Nombre=" & sNombren
    tabla= ofv.GenTabla("","aaa","","80%","","0","","-1","2","","",fila,"si")
    Formulario = ofv.GenForm("Subconsulta","",sAccion,"","post","",tabla,"si")

  sHTML = sHTML & "<center>"
  sHTML = sHTML & ofv.GenTabla("","","","","","0","","7","0","","",Formulario,"si")
  sHTML = sHTML & "</center>"

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","","")

  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconn(cn)
end if

function selecEmpre()

sImagenx = "<B><FONT FACE=CorFlatIcon COLOR=#aa0000 SIZE=4>x</FONT></B>"

   IF  Session("EMPMUN") = "S" THEN

  ' Fila    = ""

 '  SINPUT3 = "&nbsp;<A class=tt style='color:#000000;' title='SELECCIONE EMPRESA' HREF='javascript:void(0);' onclick='shwDem();' >EMPRESA:&nbsp;&nbsp;</a>"

 '  Fila    = ""
 '  Celda   = ""
  ' sHref   = ""

  ' sImage  = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
  ' sLink   = ofv.GenLink("","ttt",sHref,"ftemas","","",sImage,"si")
   'Celda   = Celda & ofv.GenCelda("","tt ","","8%","","","",SINPUT3,"si")
  ' Celda   = Celda & ofv.GenCelda("","ttt ","center","5%","","","",sLink,"si")
  ' Celda   = Celda & ofv.GenCelda("","ttt  ","","","","","","&nbsp;<BIG>" & ATREMP(0,4) & "</BIG>","si")

  ' Fila    = fila & ofv.GenRow("","","top","","","","",Celda,"si")

  ' Tabla   = ofv.GenTabla("","aaa style='"  & Session("CBRDABL") & "' ","","100% ","","0","","-1","-2","","",Fila,"si")





  ' shtml=shtml & "<script language=javascript>" & chr(13)
  ' shtml=shtml & "function shwDem(){" & chr(13)
  ' shtml=shtml & "document.all.DEM1.style.display='block';" & chr(13)
  ' shtml=shtml & "}" & chr(13)
  ' shtml=shtml & "</script>" & chr(13)
 '----------------------------------------------
  ' sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
  ' sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
  ' sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "document.all.DEM1.style.display='none';" & chr(34)
  ' sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
  ' TABLA = Tabla & "<DIV id='DEM1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:240px;left:20px;height:250px;width:350px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"


  ' TABLAX = TablaX & OFV.gentabla("", "aaa style='filter:shadow(color=#dddddd,direction=125);' ", "", "100%", "", "0", "", "-2", "-1", "", "", "", "no")
 '  TABLAX = TablaX & "<tr><td  "
 '  TABLAX = TablaX & "class=vT align=left colspan=3 STYLE='border-style:outset;border-width:2;background-color:#eeeeee;' >" & sinput2 & "&nbsp;SELECCIONE EMPRESA</td></tr>"
 '  TABLAX = TablaX & "<tr><td colspan=3 "
 '  TABLAX = TablaX & " >&nbsp;</td></tr>"

  ' FOR IAEMP=1 TO TAEMP

    '''  RESPONSE.WRITE  ATREMP(IAEMP,2)

  '    IF ATREMP(IAEMP,2) = 1 THEN

  '     IF ATREMP(IAEMP,0) = ATREMP(0,0) THEN
   '       TABLAX = TablaX & "<tr BGCOLOR=#EEEEEE >"
   '    ELSE
   '       TABLAX = TablaX & "<tr>"
   '    END IF

   '    TABLAX = TablaX & "<td width=5 "
   '    IF ATREMP(IAEMP,0) = ATREMP(0,0) THEN
   '       TABLAX = TablaX & "class=vt  >" & sImagen4si & "</td>"
   '    ELSE
   '       TABLAX = TablaX & "class=vt  >" & sImagenI3 & "</td>"
  '     END IF

 '      sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(IAEMP,5),"0","")
  '     TABLAX = TablaX & "<td width=5 "
  '     TABLAX = TablaX & "class=vt  >" & sImage & "</td>"

  '     TABLAX = TablaX & "<td class=vt  >"

       'sHref = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0&EMPRE=" & ATREMP(IAEMP,0)
 	'	sHref = "SubconsultaComponentes.asp?EMPRE=" & ATREMP(IAEMP,0) & "&Nombre=" & sNombren & "&Menu=" & sMenu
     '''  sHref = "OBJtreeTema.asp?Tipo=" & sTipo & "&Nombre=" & LNombre & "&Menu=" & SESSION("MENU") & "&EMPRE=" & ATREMP(IAEMP,0)
   '    sLink = ofv.GenLink("","cc",sHref,"","","","<b>" & ATREMP(IAEMP,4) & "</b>","si")
    '   TABLAX = TablaX & sLink
    '   TABLAX = TablaX & "</td>"
    '   TABLAX = TablaX & "</tr>"
    '  END IF

   'NEXT
'
  ' TABLAX = TablaX & OFV.genFtabla()

  ' TABLA = Tabla & TABLAX

  ' TABLA = Tabla & "</DIV>"
   TableEmp=GenEmpresas(ATREMP,"CONSCMP","")
   sHTML = sHTML & TableEmp

  end if

end function

response.write sHTML

%>