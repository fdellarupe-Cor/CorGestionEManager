<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  IDTXT = request.querystring("IDTXT")
  vuelvo = request.querystring("vuelve")
  vuelvo = replace(vuelvo,"codobjeto","Codobjeto")
  vuelve = replace(vuelvo,"***","_")
  volver = "_Vuelve=" & replace(vuelve,"_","***")
  addqry = replace(vuelve,"_","&")

  nblk = request.querystring("nblk")
  if len(nblk) = 0 then nblk = "D"

  ord = request.querystring("ord")
  if len(ord) = 0 then ord = "T"

  newdir = request.querystring("newdir")
  if len(newdir) = 0 then newdir = "0"

  opc = request.querystring("opc")




  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
'modificacion 14/6
 if session("EMPMUN") = "S" then
 Draiz="Emp" & SESSION("D" & nCodigo & "-0" ) & "/estructura"
 else
 Draiz= "estructura"
 end if


  esttxt = " STYLE={font-family:tahoma;color:#000000;font-size:10pt;background-color:#f7eedd;} "

        ccomp = 0
        tcomp =  sTipo

  set cn = ofv.conectar(ofv.strconn2)
  opc = "NO"

  strsql = "SELECT * from ObjPORTADA where objid = " & IDTXT

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     if opc = "NO" then
        ccomp = rs(9)
        ccomp2 = rs(9)
        if session("EMPMUN") = "S" then
		           docanterior=Draiz & "/" & rs(9)
		        else
		           docanterior=   "estructura/" & rs(9)
                end if
     end if
  end if


'  RESPONSE.WRITE "ccomp " & ccomp & " ccomp2 " & ccomp2 & " <br>"
  ccompS = request.form("comp1")
  if nblk = "S" then ccompS = request.form("comp1")
  if nblk = "S" then newdir = ccompS
  if nblk = "T" then ccompT = request.form("comp2")
  if nblk = "T" then ccomp2 = ccompT

  if nblk = "D" then ccomp2 = ccomp

  if nblk = "D" and ord = "F" then
     session("comp31") = request.form("comp31")
     session("comp32") = request.form("comp32")
  end if

  if session("comp31") = "" then
     session("comp31") = "TODOS"
  end if


  if session("comp32") = "" then
     session("comp32") = "JPG"
  end if


     comp31 = session("comp31")
     comp32 = session("comp32")



  Set fso = CreateObject("Scripting.FileSystemObject")

 '''' sHTML = ofv.menuheader("","")


  sHTML = ofv.FormHeader(session("FormN"))
' shtml=replace(shtml,"<body","<body onmousemove=" &chr(34) &"aaaxxx(event);"&chr(34))



  'sHTMLf =  " <script >"
  'sHTMLf = sHTMLf & " function aaaxxx(e){ "
  'sHTMLf = sHTMLf & "var obj=e.srcElement;"
  ''sHTMLf = sHTMLf &"if (obj){"
  ''sHTMLf = sHTMLf &"window.status=obj.id + '----' + obj.innerHTML;"
  'sHTMLf = sHTMLf &"}"

'sHTMLf = sHTMLf &"</script>"
'shtml=replace(shtml,"</HEAD>", sHTMLf & "</head>" )
'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------



  sHTML = sHTML & "<center><br>"

  sHTML = sHTML & ofv.GenTabla("panel","aaa ","","80%","",0,0,0,0,"","","","no")

      sicK = "<font face='corflaticon' size=4 color=#993333 >&#0163;4</font>&nbsp;"
     Celda =  sicK & "<b>Modificacion de Imagen de contenido</b>"

'  sHTML = sHTML & "<tr>"
'  sHTML = sHTML & "<td width=50%  colspan=2 class=tt   >"
'  sHTML = sHTML & Celda
'  sHTML = sHTML & "</td>"
'  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td width=50%  class=ttt align=center  >"
  sHTML = sHTML & "Directorios&nbsp;disponibles"
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td width=50%  class=ttt align=center  >"
  sHTML = sHTML & "Imagenes&nbsp;disponibles"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td width=50% align=center>"




  sPath = Request.ServerVariables("PATH_INFO")

  Xruta = Request.ServerVariables("SERVER_NAME") & left(sPath,instrRev(sPath,"/",-1) - 1)
' response.write Xruta & " 2<br>"
 sPath = server.mappath(sPath)
'  response.write spath & " 3<br>"
    sPath1 = replace(sPath,"edestrdetIMG.asp","",1,-1,1) & "ESTRUCTURA"
'  response.write spath & " 4<br>"
'  Set Folder = fso.getfolder(sPath1)
 ' set sfolder = Folder.SubFolders
'  Set Files = Folder.Files

  if bModificar then
      sAccion = "edPORTADAIMG.asp?"
      sAccion = sAccion & "nblk=S&ord=" & ord & "&Codobjeto=" & nCodigo & "&CantRegAMover=" & CantRegAMover
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&opc=" & opc & "&IDTXT=" & IDTXT & "&newdir=" & newdir & "&vuelve=" & vuelvo
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action="  & chr(34) &  sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "<font class=at><select  size=10   name=comp1 style={border-style:outset;border-width:4;} "
  sHTML = sHTML & " onchange=" & chr(34) & "fcomp1.submit();" & chr(34)

  sHTML = sHTML & " >"

   if ccompS = ""  then
    'response.write " 222"
       'sHTML = sHTML & " selected "
       sHTML = sHTML & "<option STYLE='background-color:#F7AA22;font-weight:bold;' value=" & Chr(34) & "" & Chr(34)
  	  sHTML = sHTML & "  "
  	 sHTML = sHTML & " >" & Draiz &"</option>"
    else
   '  response.write ccompS
     if InStrRev(ccompS & "" ,"/") = 0  then
       sHTML = sHTML & "<option STYLE='background-color:#F7AA22;font-weight:bold;' value=" & Chr(34) & "" & Chr(34)
     else
       sHTML = sHTML & "<option STYLE='background-color:#F7AA22;font-weight:bold;' value=" & Chr(34) & Left(ccompS,InStrRev(ccompS,"/",-1) - 1) & Chr(34)
     end if
    	 sHTML = sHTML & "  "
    	' response.write Draiz
  	 sHTML = sHTML & " >" & Draiz  &"/" & ccompS & "</option>"
  end if

  sfolder=""


  '  response.write Draiz & "/" & ccompS

   if trim(cstr(ccompS )) <> "" then

    sfolder=split(getfolders(Draiz & "/" & ccompS,""),"|")
    else
    sfolder=split(getfolders(Draiz,""),"|")
  end if



for ft=0 to ubound(sfolder)-1
     if ccompS="" then
       sHTML = sHTML & "<option value=" & chr(34)  & sfolder(ft)   & chr(34)
     else
       sHTML = sHTML & "<option value=" & chr(34)  & ccompS & "/" & sfolder(ft)   & chr(34)
     end if

     'if ucase(newdir) = ucase(("/" & sfolder(ft))) then   sHTML = sHTML & " selected "
     if ccompS="" then
     sHTML = sHTML & " >" & Draiz & "/"   & Ucase(sfolder(ft)) & "</option>"
     else
     sHTML = sHTML & " >" & Draiz & "/" & ccompS  & "/" & Ucase(sfolder(ft)) & "</option>"
     end if

  Next




  sHTML = sHTML & "</select></font>"
  sHTML = sHTML & "</FORM>" & chr(13)

  sHTML = sHTML & "</td>"
  sHTML = sHTML & "<td width=50% align=center  >"
'-------------
  if bModificar then
      sAccion = "edPORTADAIMG.asp?"
      sAccion = sAccion & "nblk=T&ord=" & ord & "&Codobjeto=" & nCodigo & "&CantRegAMover=" & CantRegAMover
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&opc=" & opc & "&IDTXT=" & IDTXT & "&newdir=" & newdir & "&vuelve=" & vuelvo
  else
     sAccion = ""
  end if

    if newdir = "0" then
       if session("EMPMUN") = "S" then
       		newdirn = "Emp" & SESSION("D" & nCodigo & "-0" ) & "/ESTRUCTURA"
       else
			newdirn = "ESTRUCTURA"
       end if
    else
       if session("EMPMUN") = "S" then
		    newdirn = "Emp" & SESSION("D" & nCodigo & "-0" ) & "/ESTRUCTURA" & newdir
	   else
			newdirn = "ESTRUCTURA" & newdir
       end if
    end if

    sPath2 = replace(sPath,"edPORTADAIMG.asp","",1,-1,1) & newdirn
'  response.write spath2 & " 4<br>"
 ' Set Folder = fso.getfolder(sPath2)
 ' Set Files = Folder.Files


  sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action="  & chr(34) &  sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "<font class=at><select size=10  name=comp2  "
  sHTML = sHTML & " onchange=" & chr(34) & "fcomp2.submit();" & chr(34)
  sHTML = sHTML & " >"



  sfcount = 0

      if comp32 = "TODOS" then
         if comp31 = "TODOS" then
            Sfiles = "SI"
         else
            Sfiles = "NO"
         end if
      else
         Sfiles = "NO"
      end if


  sfilesn=split(getfilesinfolder(Draiz & "/" & ccompS,""),"|")
  'For Each F in Files
     for f=0 to ubound(sfilesn)-1

      file=split(sfilesn(f),"###")

      if Sfiles = "NO" then
         if comp32 = "TODOS" then
            sfilesva1 = "SI"
         else
            sfext = instr(file(0),".")
            if sfext = 0 then
               sfext = "NO"
            else
               sfext = ucase(right(file(0),len(file(0))-sfext))
            end if
            if sfext = comp32 then
               sfilesva1 = "SI"
            else
               sfilesva1 = "NO"
            end if
         end if

         if comp31 = "TODOS" then
            sfilesva2 = "SI"
         else
            sfextn = len(comp31)
            sfext = ucase(left(file(0),sfextn))
            if sfext = Ucase(comp31) then
               sfilesva2 = "SI"
            else
               sfilesva2 = "NO"
            end if
         end if

'      response.write "sfiles " & sfiles & " comp32 " & comp32 & " comp31 " & comp31 & " va1 " & sfilesva1 & " va2 " & sfilesva2 & " name " & f.name & " <br>"

         if sfilesva1 = "SI" and sfilesva2 = "SI" then
 '           response.write " entro " & f.name & " <br>"
            sfilesva = "SI"
         else
            sfilesva = "NO"
         end if
      else
         sfilesva = "SI"
      end if



      if sfilesva = "SI" then
         sfcount = sfcount + 1
         if newdir = "0" then
              sHTML = sHTML & "<option value=" & chr(34) & file(0) & chr(34)
         else
              sHTML = sHTML & "<option value=" & chr(34)  & file(0) & chr(34)
         end if

             if ccomp2 =   file(0) then   sHTML = sHTML & " selected "
              sHTML = sHTML & " >" & Ucase( file(0)) & "</option>"

      end if

  Next

  if sfcount = 0 then
              sHTML = sHTML & "<option value=" & chr(34) & ccomp & chr(34)
              sHTML = sHTML & " selected "
              sHTML = sHTML & " >No hay imagenes disponibles</option>"

   end if


  sHTML = sHTML & "</select></font>"
   sHTML = sHTML & "<INPUT TYPE=hidden NAME=comp1 VALUE=" & CHR(34)& ccompS & CHR(34) & "></INPUT>"
  sHTML = sHTML & "</FORM>" & chr(13)
'--------
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"


  sHTML = sHTML & "<tr>"

  sHTML = sHTML & "<td colspan=2 align=center  >"
'-------------
  if bModificar then
      sAccion = "edPORTADAIMG.asp?"
      sAccion = sAccion & "nblk=D&ord=F&Codobjeto=" & nCodigo & "&CantRegAMover=" & CantRegAMover
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&opc=" & opc & "&IDTXT=" & IDTXT & "&newdir=" & newdir & "&vuelve=" & vuelvo

  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp3 ID=fcomp3 action="  & chr(34) &  sAccion  & chr(34) &  "  method=post >" & chr(13)

  sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=3  ","","100%","",0,0,0,0,"","","","no")
  sHTML = sHTML & "<tr>"

 ' sHTML = sHTML & "<td width=15% class=ttt   >"
 ' sHTML = sHTML & "FILTRO: "
 ' sHTML = sHTML & "</td>"


  sHTML = sHTML & "<td width=35% align=center  >"
  sHTML = sHTML & ofv.GenerarInput("comp31",comp31,"agregar","","80","dt","")
  sHTML = sHTML & "</td>"


  sHTML = sHTML & "<td width=35% align=center  >"

  sHTML = sHTML & "<font class=at><select  name=comp32  "
  sHTML = sHTML & " >"

  sHTML = sHTML & "<option value=" & comp32
  sHTML = sHTML & " selected "
  sHTML = sHTML & " >" & comp32 & "</option>"

  sHTML = sHTML & "<option value=TODOS "
  sHTML = sHTML & " >TODOS</option>"
  sHTML = sHTML & "<option value=JPG "
  sHTML = sHTML & " >JPG</option>"
  sHTML = sHTML & "<option value=GIF "
  sHTML = sHTML & " >GIF</option>"
  sHTML = sHTML & "<option value=BMP "
  sHTML = sHTML & " >BMP</option>"



  sHTML = sHTML & "</select></font>"
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td width=30% align=center >"
  sHTML = sHTML & "<INPUT TYPE=HIDDEN NAME=comp1 VALUE=" & CHR(34)& ccompS & CHR(34) & "></INPUT>"
  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='FILTRAR' class=bt >"
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "</tr>"
  sHTML = sHTML & "</table>"
  sHTML = sHTML & "</FORM>" & chr(13)
'--------



  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"


  sHTML = sHTML & "</table>"

  sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=2 style='border-style:solid;border-width:1;border-color:" & session("cmfmenu") & ";' ","","80%","",0,0,0,0,"","","","no")
  sHTML = sHTML & "<tr>"

   opc = "NO"
if opc = "NO" then
  sHTML = sHTML & "<td align=center valign=top style='border-style:solid;border-width:1;border-color:#dddddd;' >"
  sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=3  ","","100%","",0,0,0,0,"","","","no")



  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  class=ttt align=center  >"
  sHTML = sHTML & "Documento&nbsp;anterior"
  sHTML = sHTML & "</td>"



  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td   class=ut   >"


if session("EMPMUN") = "S" then
     ncomp = "Emp" & SESSION("D" & nCodigo & "-0" ) & "/ESTRUCTURA/" & ccomp
else
	 ncomp = "ESTRUCTURA/" & ccomp
end if

  sHTML = sHTML & "<font color=#000080 ><B>" & docanterior & "</b></font>"
  sHTML = sHTML & "</td>"


  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
if session("EMPMUN") = "S" then
  sImage = ofv.GenImage("Imagen","aaa width=100 border=1 ","","",docanterior,"","")
else
  sImage = ofv.GenImage("Imagen","aaa width=100 border=1 ","","",docanterior,"","")
end if

  sHTML = sHTML & ofv.GenCelda("aaa ","","center","","","","",sImage,"si")
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "</table>"
  sHTML = sHTML & "</td>"

  xcolsp = ""

else
  xcolsp = " colspan=2 "

end if
  sHTML = sHTML & "<td " & xcolsp & " align=center valign=top style='border-style:solid;border-width:1;border-color:#dddddd;' >"
  sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=3  ","","100%","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  class=ttt align=center  >"
  sHTML = sHTML & "Documento&nbsp;nuevo"
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "</tr>"

'  if nblk = "D" then ccomp2 = ccomp
'  if nblk = "S" then ccomp2 = ccompS
'  if nblk = "T" then ccomp2 = ccompT

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  class=ut   >"

if session("EMPMUN") = "S" then
     ncomp = draiz & "/" & ccompS & "/" & ccomp2
else
	 ncomp = "ESTRUCTURA/" & ccomp2
end if

  sHTML = sHTML & "<font color=#993333 ><B>" & ncomp & "</b></font>"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
if session("EMPMUN") = "S" then
  sImage = ofv.GenImage("Imagen","aaa width=100 border=1 ","","", ncomp,"","")
else
  sImage = ofv.GenImage("Imagen","aaa width=100 border=1 ","","","ESTRUCTURA/" & ncomp,"","")
end if
  sHTML = sHTML & ofv.GenCelda("aaa","","center","","","","",sImage,"si")
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "</table>"
  sHTML = sHTML & "</td>"

'''''''''''''''''''''

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td align=center  >"
  if bModificar then

      sAccion = chr(34) & "asp/ActualizarEN.asp?ID=" & idtxt & "&Tabla=ObjPORTADA"
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&volver=../edPORTADAdet.asp?" & vuelvo & chr(34)



  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=acep ID=acep action=" & sAccion & "  method=post >" & chr(13)
  sHTML = sHTML & ofv.GenerarInput(rs(9).name,ccompS & "/" & ccomp2,"hidden","","100","","")
  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='ACEPTAR' class=bt >"
  sHTML = sHTML & "</FORM>" & chr(13)
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td  align=center   >"

    sAccion = "edPORTADAdet.asp?" & addqry
  sHTML = sHTML & "<FORM  name=Cancel ID=Cancel action=" & sAccion & "  method=post >" & chr(13)
  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='CANCELAR' class=bt >"
  sHTML = sHTML & "</FORM>" & chr(13)
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "</tr>"
  sHTML = sHTML & "</table>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</center>"


'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------



  sHTML = sHTML & "</body></html>"

end if

response.write sHTML

%>