<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
'response.write  request.querystring
  ofv.ObtenerAtributos Session("Form"),""
  bModificar = ofv.AModificar

  CantRegAMover = request.querystring("CantRegAMover")
  nblk = request.querystring("nblk")
  if  len(nblk) = 0 then nblk = "D"

  ord  = request.querystring("ord")
  if  len(ord) = 0 then ord = "T"

  newdir = request.querystring("newdir")
  if  len(newdir) = 0 then newdir = "0"

  opc  = request.querystring("opc")
  cid  = request.querystring("cid")
  xp=request.querystring("xp")

   'Draiz="Emp" & SESSION("un" ) & "/ARCHIVOS/IMAGENES"

   if session("EMPMUN") = "S" then
   	    Draiz="Emp" & SESSION("un" ) & "/ARCHIVOS/IMAGENES"
   else
   		Draiz= "ARCHIVOS/IMAGENES"
   end if

  nCodigo = request.querystring("Codobjeto")
  sTipo   = request.querystring("Tipobjeto")
  sfamily = request.querystring("family")
  idx   = request.querystring("idx")
  vueltaC = request.querystring("volver")
  seditar=request.querystring("editar")
  volver2 = ofv.vueltaasp2("../")
'response.write request.querystring
  if nCodigo<>"" then
   session("nXCodigo")=nCodigo
  else
   nCodigo=session("nXCodigo")
  end if

  if IDX <> "" then
     session("idx")=idx
    else
     idx=session("idx")
  end if
  if session("sfamily")="" then
     session("sfamily")=sfamily
    else
     sfamily=session("sfamily")
  end if

  esttxt = " STYLE={font-family:tahoma;color:#000000;font-size:10pt;background-color:#f7eedd;} "

  ccomp  = 0
  tcomp  =  sTipo

  set cn = ofv.conectar(ofv.strconn2)
  xentro="N"
  Strsql = "Select * From novegenparimg "
 if opc = "NO" then
     strsql = strsql & " inner join novegenpar on (novegenpar.objid=novegenparimg.novegenpar) "
     strsql = strsql & " Where  novegenparimg.OBJID =" & IDX
 end if
'response.write request.form
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if  not rs.eof then
      if  opc = "NO" then
          rs5=rs(5)&""
          rs5=replace(rs(5),"//","/")
          ccomp = rs5
          ccomp2 = rs5
          XENTRO="S"
          emp=rs("codemp")
          IF CCOMP="" THEN
		    XDIRP="/"
		  ELSE
		    xdirp= Left(ccomp,InStrRev(ccomp,"/",-1) - 1)
		 END IF
         DOCANTERIOR=ccomp

      end if
  end if




  if  nblk = "S" then ccompS = request.form("comp1")
  if  nblk = "S" then newdir = ccompS
  if  nblk = "T" then ccompT = request.form("comp2")
  if  nblk = "T" then ccomp2 = ccompT

  if  nblk = "D" then ccomp2 = ccomp

  if  nblk = "D" and ord = "F" then
      session("comp31") = request.form("comp31")
      session("comp32") = request.form("comp32")
  end if

  if  session("comp31") = "" then
      session("comp31") = "TODOS"
  end if


  if  session("comp32") = "" then
      session("comp32") = "JPG"
  end if


  comp31 = session("comp31")
  comp32 = session("comp32")

  xdir=request.form("comp1")

     IF XDIR <> "RAIZ" AND XDIR <> "" THEN
     xdirp=xdir
     ELSE
     XDIRP="/"
     END IF



  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  'shtml = shtml & CorPnlFnc("N","")

  'sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------



  Set fso = CreateObject("Scripting.FileSystemObject")

  sHTML = sHTML & "<center><br>"



  sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3 style={border-style:solid;border-width:1;border-color:#000000;} ","","80%","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td width=50%  class=tt align=center  >"
  sHTML = sHTML & "Directorios&nbsp;disponibles"
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td width=50%  class=tt align=center  >"
  sHTML = sHTML & "Imagenes&nbsp;disponibles"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td width=50% align=center  >"

'  response.write direct & " 0<br>"


  sPath = Request.ServerVariables("PATH_INFO")
'  response.write spath & " 1<br>"
  Xruta = Request.ServerVariables("SERVER_NAME") & left(sPath,instrRev(sPath,"/",-1) - 1)
' response.write Xruta & " 2<br>"
 sPath = server.mappath(sPath)
'  response.write spath & " 3<br>"
    sPath1 = replace(sPath,"EdimagenCMPex.asp","",1,-1,1) & "archivos"
'  response.write spath & " 4<br>"
'  Set Folder = fso.getfolder(sPath1)
'  set sfolder = Folder.SubFolders
'  Set Files = Folder.Files
'response.write request.form
  if bModificar then
      sAccion = "EdimagenCMPexNOVED.asp?"
      sAccion = sAccion & "xp=N&nblk=S&ord=" & ord & "&Codobjeto=" & nCodigo & "&CantRegAMover=" & CantRegAMover
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&opc=" & opc & "&cid=" & cid & "&newdir=" & newdir
       sAccion = sAccion & "&editar=" & seditar
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action="  & chr(34) &  sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "<font class=at><select size=10  name=comp1 style={border-style:outset;border-width:4;} "
  sHTML = sHTML & " onchange=" & chr(34) & "fcomp1.submit();" & chr(34)

  sHTML = sHTML & " >"

  if Left(xdirp,InStrRev(xdirp,"/",-1) - 1)= ""  then

       sHTML = sHTML & "<option STYLE='background-color:" & Session("CMFPNBOT") & ";font-weight:bold;' value=" & Chr(34) & "RAIZ" & Chr(34)
       sHTML = sHTML & "  "
  	 sHTML = sHTML & " >" & draiz & XDIRP &"</option>"
      else
  	 sHTML = sHTML & "<option STYLE='background-color:#F7AA22;font-weight:bold;' value=" & Chr(34) & Left(xdirp,InStrRev(xdirp,"/",-1) - 1) & Chr(34)
  	 sHTML = sHTML & "  "
  	 sHTML = sHTML & " >" & draiz & xdirp & "</option>"
      end if



  	If session("EMPMUN") = "S" Then

    '   RESPONSE.WRITE xdirp

  		sfolder=Split(getfolders(draiz & xdirp,""),"|")


  	Else
  		'sHTML = sHTML & "ESTRUCTURA</option>"
  		'sfolder=Split(getfolders("ARCHIVOS/IMAGENES",""),"|")
  		 sfolder=Split(getfolders(draiz & xdirp,""),"|")
	End If
  For f=0 To UBound(sfolder)-1
      '  response.write XDIRP & "<br>"
  	    IF XDIRP="/" THEN XDIRP=""
  		sHTML = sHTML & "<option value=" & Chr(34) & xdirp &"/" & sfolder(f) & Chr(34)
  		If newdir = ("/" & sfolder(f)) Then   sHTML = sHTML & " selected "
  		If session("EMPMUN") = "S" Then
  			sHTML = sHTML & " >" & draiz & xdirp &"/" & UCase(sfolder(f)) & "</option>"
  		Else
  			sHTML = sHTML & " >ARCHIVOS/IMAGENES/" & UCase(sfolder(f)) & "</option>"
  		End If





	Next




  sHTML = sHTML & "</select></font>"
  sHTML = sHTML & "</FORM>" & chr(13)

  sHTML = sHTML & "</td>"
  sHTML = sHTML & "<td width=50% align=center  >"
'-------------
  if bModificar then
      sAccion = "EdimagenCMPexNOVED.asp?"
      sAccion = sAccion & "nblk=T&ord=" & ord & "&Codobjeto=" & nCodigo & "&CantRegAMover=" & CantRegAMover
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&opc=" & opc & "&cid=" & cid & "&newdir=" & newdir
      sAccion = sAccion & "&editar=" & seditar
  else
     sAccion = ""
  end if

   ' if newdir = "0" then
   '    newdirn = "Emp" & emp & "/archivos/imagenes"
   ' else
   '    newdirn = "Emp" & emp & "/archivos/imagenes" & newdir
   ' end if

    sPath2 = replace(sPath,"EdimagenCMPexNOVED.asp","",1,-1,1) & newdirn
'  response.write spath2 & " 4<br>"
  'Set Folder = fso.getfolder(sPath2)
 ' Set Files = Folder.Files
'response.write DRAIZ & xdirp
  posdir=DRAIZ & xdirp
  sfilesn=split(getfilesinfolder(DRAIZ & xdirp,""),"|")
  sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action="  & chr(34) &  sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "<font class=at><select size=10  name=comp2 style={border-style:outset;border-width:4;} "
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
              'sHTML = sHTML & "<option value=" & chr(34) & replace(newdir,"/","",1) & "/" & file(0) & chr(34)
              sHTML = sHTML & "<option value=" & chr(34)  & file(0) & chr(34)
         end if

            ' if ccomp2 = (newdir & "/" & file(0)) then   sHTML = sHTML & " selected "
            '  sHTML = sHTML & " >" & Ucase(newdirn & "/" & file(0)) & "</option>"
             if ccomp2 =  file(0) then   sHTML = sHTML & " selected "
		  sHTML = sHTML & " >" & Ucase( file(0)) & "</option>"


      end if

  Next

  if sfcount = 0 then
              sHTML = sHTML & "<option value=" & chr(34) & ccomp & chr(34)
              sHTML = sHTML & " selected "
              sHTML = sHTML & " >No hay imagenes disponibles</option>"

   end if


  sHTML = sHTML & "</select></font>"
  sHTML = sHTML & "<INPUT TYPE=HIDDEN NAME=comp1 VALUE=" & CHR(34)& XDIRP & CHR(34) & "></INPUT>"
  sHTML = sHTML & "</FORM>" & chr(13)
'--------
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"


  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td width=50% align=center class=tt   >"
  sHTML = sHTML & "&nbsp;"
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td width=50% align=center class=tt   >"
  sHTML = sHTML & "FILTRO: "
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"


  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td width=50% align=center  >"
  sHTML = sHTML & "&nbsp;"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "<td width=50% align=center  >"
'-------------
  sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=3 style={border-style:solid;border-width:1;border-color:#000000;} ","","100%","",0,0,0,0,"","","","no")
  sHTML = sHTML & "<tr>"

  if bModificar then
      sAccion = "EdimagenCMPexNOVED.asp?"
      sAccion = sAccion & "xp=N&nblk=D&ord=F&Codobjeto=" & nCodigo & "&CantRegAMover=" & CantRegAMover
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&opc=" & opc & "&cid=" & cid & "&newdir=" & newdir
       sAccion = sAccion & "&editar=" & seditar

  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp3 ID=fcomp3 action="  & chr(34) &  sAccion  & chr(34) &  "  method=post >" & chr(13)

  sHTML = sHTML & "<td width=35% align=center  >"
  sHTML = sHTML & ofv.GenerarInput("comp31",comp31,"agregar","","80","dt","")
  sHTML = sHTML & "</td>"


  sHTML = sHTML & "<td width=35% align=center  >"

  sHTML = sHTML & "<font class=at><select  name=comp32  "
 ' sHTML = sHTML & " onchange=" & chr(34) & "fcomp3.submit();" & chr(34)

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

  sHTML = sHTML & "<td width=30% align=center  >"
  sHTML = sHTML & "<INPUT TYPE=HIDDEN NAME=comp1 VALUE=" & CHR(34)& XDIRP & CHR(34) & "></INPUT>"
  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='FILTRAR' class=bt >"
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "</FORM>" & chr(13)

  sHTML = sHTML & "</tr>"
  sHTML = sHTML & "</table>"
'--------
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"


  sHTML = sHTML & "</table>"

  sHTML = sHTML & ofv.GenTabla("panel2 ","aaa cellpadding=3 style={border-style:solid;border-width:1;border-color:#000000;} ","","80%","",0,0,0,0,"","","","no")
  sHTML = sHTML & "<tr>"

if opc = "NO" then
  sHTML = sHTML & "<td valign=top width=50% >"
  sHTML = sHTML & ofv.GenTabla("panel2 ","aaa cellpadding=3  ","","100%","",0,0,0,0,"","","","no")



  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  class=tt align=center >"
  sHTML = sHTML & "Imagen&nbsp;anterior"
  sHTML = sHTML & "</td>"



  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td    class=cc   >"



     ncomp = "" & ccomp

  sHTML = sHTML & "<font color=#000080 ><B>" & DOCANTERIOR & "</b></font>"
  sHTML = sHTML & "</td>"


  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"

  sImage = ofv.GenImage("Imagen","aaa width=100 border=1 ","","","" & DOCANTERIOR,"","")
  sHTML = sHTML & ofv.GenCelda("aaa ","","center","","","","",sImage,"si")




  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "</table>"
  sHTML = sHTML & "</td>"

  xcolsp = ""

else
  xcolsp = " colspan=2 "

end if
  sHTML = sHTML & "<td " & xcolsp & "   valign=top width=50% >"
  sHTML = sHTML & ofv.GenTabla("panel2 ","aaa cellpadding=3  ","","100%","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  class=tt align=center  >"
  sHTML = sHTML & "Imagen&nbsp;nueva"
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "</tr>"

'  if nblk = "D" then ccomp2 = ccomp
'  if nblk = "S" then ccomp2 = ccompS
'  if nblk = "T" then ccomp2 = ccompT

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  class=cc   >"


if xp="N" or xp="q" then
 ncomp = ccomp2
else

 ncomp = NEWDIRN &"/" & ccomp2
end if
'response.write ncomp
'response.write xp
if ccompT <> "" then
  sHTML = sHTML & "<font color=#993333 ><B>"    & ccomp2 & "</b></font>"
else
sHTML = sHTML & "<font color=#993333 ><B>&nbsp;</b></font>"
end if

  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"
'response.write request.querystring
  sHTML = sHTML & "<tr>"
  If session("EMPMUN") = "S" Then
  		sImage = ofv.GenImage("Imagen","aaa width=100 border=1 ","","", Draiz & XDIRP  & "/" &ccomp2,"","")
  	Else
  		'sImage = ofv.GenImage("Imagen","aaa width=100 border=1 ","","","ESTRUCTURA/" & ccomp2,"","")
  		 if ccompT <> "" then
  	      sImage = ofv.GenImage("Imagen","aaa width=100 border=1 ","","", Draiz &  XDIRP  & "/" &ccomp2,"","")
  	     else
  	      sImage=""
  	     end if
	End If
 ' sImage = ofv.GenImage("Imagen","aaa width=100 border=1 ","","","" & ncomp,"","")
  sHTML = sHTML & ofv.GenCelda("aaa  ","","center","","","","",sImage ,"si") '& ofv.GenerarInput(rs(5).name,newdirn &"/" & ccomp2,"hiddqen","","10","","")
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "</table>"
  sHTML = sHTML & "</td>"

'''''''''''''''''''''



  if bModificar then

    if opc = "NO" then
      ' saccion ="Adnovedadesimg.asp?"
      'sAccion = sAccion & "Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo & "&family=" & sfamily & "&idx=" & idx & "&opcion=UPD"
       saccion ="Adnovedadesimg.asp?id=" & nCodigo & "&opcion=UPD" &"&idx=" & idx
        sAccion = sAccion & "&editar=" & seditar
    '  response.write saccion
    else

       'sAccion = "Adnovedadesimg.asp"
        saccion ="Adnovedadesimg.asp?id=" & nCodigo & "&opcion=UPD" &"&idx=" & idx
         sAccion = sAccion & "&editar=" & seditar
       'sAccion = sAccion & "Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo & "&family=" & sfamily & "&idx=" & idx & "&opcion=UPD"

    end if


  else
     sAccion = ""
  end if


  sAccionac = sAccion

  sHTML = sHTML & ofv.GenForm("AgregarControlb","",sAccionac,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panelxx","aaa cellpadding=3  ","","80%","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<tr>"
 ' sImage = ofv.GenerarInput(rs(5).name,newdirn & "/" & ccomp2,"hidden","","10","","")
  sImage = ofv.GenerarInput(rs(5).name, Draiz & "/" & XDIRP  & "/" & ccomp2,"HIDDEN","","100","","")
  sHTML = sHTML & ofv.GenCelda("aaa","","center","","","","",sImage,"si")
  sHTML = sHTML & "</tr>"



'  sInput  = ofv.GenerarInput("","Aceptar","submit","","10","bt","") & ofv.GenerarInput(rs(3).name,ccomp2,"hidden","","10","","")
'  Celda   = ofv.GenCelda("","","","","","","",sInput,"si")
'  Agregar = ofv.GenForm("Agregarb","",sAccion,"","post","",Celda,"si")

  'sAccion = "edImagen2.asp?Codcomp=" & nCodigo & "&Opcion=no&CantRegAMover=" & CantRegAMover
  'sAccion = sAccion & "&Tipocomp=" & sTipo

'  sInput   = ofv.GenerarInput("","Cancelar","submit","","10","bt","")
'  Celda    = ofv.GenCelda("","","","","","","",sInput,"si")
'  Cancelar = ofv.GenForm("AgregarC","",sAccion,"","post","",Celda,"si")



  sHTML = sHTML & "</table>"


  'Botones Aceptar y Cancelar
  ' sAccion = "edgeneralinfoimg.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
   saccion ="Adnovedadesimg.asp?id=" & nCodigo &"&idx=" & idx
   'sAccion = sAccion & "Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo & "&family=" & sfamily & "&idx=" & idx & "&opcion=NO"
  sHTML = sHTML & ofv.AgregarBoton("AgregarControlb",sAccion,"_self")


  sHTML = sHTML & "</table>"


  sHTML = sHTML & "</center>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


 ' Buscar  = Cancelar
 ' Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

'--------------------------------------------------------------------------------
 ' sHTML = sHTML & "</div>"
'
 ' Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
 ' Fila = ofv.GenRow("","","","","","","",Celda,"si")
 ' sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

 ' sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------


  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>