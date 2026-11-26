<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form"),""
  bModificar = ofv.AModificar



  sDocum = request.querystring("itemdata")


  vuelta = ofv.vueltaasp2("../") &  "?itemdata=" & sDocum




  nblk = request.querystring("nblk")
  if len(nblk) = 0 then nblk = "D"

  ord = request.querystring("ord")
  if len(ord) = 0 then ord = "T"

  newdir = request.querystring("newdir")
  if len(newdir) = 0 then newdir = "0"

  opc = request.querystring("opc")





  vueltaC = request.querystring("volver")
  volver2 = ofv.vueltaasp2("../")

  esttxt = " STYLE={font-family:tahoma;color:#000000;font-size:10pt;background-color:#f7eedd;} "

        ccomp = 0
        tcomp =  sTipo

  set cn = ofv.conectar(ofv.strconn6)
  opc = "NO"

  StrSql = "Select * From PROYantec where id = " & sDocum

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     if opc = "NO" then
        ccomp = rs(3)
        ccomp2 = rs(3)
     end if  
  end if


'  RESPONSE.WRITE "ccomp " & ccomp & " ccomp2 " & ccomp2 & " <br>"

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
     session("comp32") = "HTM"
  end if


     comp31 = session("comp31") 
     comp32 = session("comp32") 



  Set fso = CreateObject("Scripting.FileSystemObject")

  sHTML = ofv.FormBody & "<center><br>"

  sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3 style={border-style:outset;border-width:2;} ","","80%","",0,0,0,0,"","","","no")



  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td width=50%  colspan=2 class=tt align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & "Descripcion:&nbsp;&nbsp;" & rs(3) & "&nbsp;-&nbsp;" & rs(4)
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td width=50%  class=tt align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & "Directorios&nbsp;disponibles"
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td width=50%  class=tt align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & "Documentos&nbsp;disponibles"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td width=50% align=center style={border-style:outset;border-width:2;} >"




  sPath = Request.ServerVariables("PATH_INFO")

  Xruta = Request.ServerVariables("SERVER_NAME") & left(sPath,instrRev(sPath,"/",-1) - 1)
' response.write Xruta & " 2<br>"
 sPath = server.mappath(sPath)
'  response.write spath & " 3<br>"
    sPath1 = replace(sPath,"PRYArcSel.asp","",1,-1,1) & "PROYECTOS/" & session("NCliente")
'  response.write spath1 & " 4<br>"
  Set Folder = fso.getfolder(sPath1)
  set sfolder = Folder.SubFolders
'  Set Files = Folder.Files

  if bModificar then
      sAccion = "PRYArcSel.asp?" 
      sAccion = sAccion & "nblk=S&ord=" & ord & "&Codobjeto=" & nCodigo & "&CantRegAMover=" & CantRegAMover
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&opc=" & opc & "&itemdata=" & sDocum & "&newdir=" & newdir
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action="  & chr(34) &  sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "<font class=at><select size=10  name=comp1 style={border-style:outset;border-width:4;} "
  sHTML = sHTML & " onchange=" & chr(34) & "fcomp1.submit();" & chr(34)

  sHTML = sHTML & " >"

  sHTML = sHTML & "<option value=" & chr(34) & "0" & chr(34)
  if newdir = "0" then 
     sHTML = sHTML & " selected " 
  end if
  sHTML = sHTML & " >" & "PROYECTOS/"  & session("NCliente") & "</option>" 

  For Each F in sfolder

              sHTML = sHTML & "<option value=" & chr(34) & "/" & f.name & chr(34)
             if newdir = ("/" & f.name) then   sHTML = sHTML & " selected "
              sHTML = sHTML & " >" & "PROYECTOS/"  & session("NCliente") & "/" & Ucase(f.name) & "</option>"


  Next




  sHTML = sHTML & "</select></font>"
  sHTML = sHTML & "</FORM>" & chr(13)

  sHTML = sHTML & "</td>"
  sHTML = sHTML & "<td width=50% align=center style={border-style:outset;border-width:2;} >"
'-------------
  if bModificar then
      sAccion = "PRYArcSel.asp?" 
      sAccion = sAccion & "nblk=T&ord=" & ord & "&Codobjeto=" & nCodigo & "&CantRegAMover=" & CantRegAMover
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&opc=" & opc & "&itemdata=" & sDocum & "&newdir=" & newdir
  else
     sAccion = ""
  end if

    if newdir = "0" then
       newdirn =  "PROYECTOS/" & session("NCliente")
    else
       newdirn =  "PROYECTOS/" & session("NCliente") & newdir
    end if

    sPath2 = replace(sPath,"PRYArcSel.asp","",1,-1,1) & newdirn
'  response.write spath2 & " 4<br>"
  Set Folder = fso.getfolder(sPath2)
  Set Files = Folder.Files


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

  For Each F in Files



      if Sfiles = "NO" then
         if comp32 = "TODOS" then    
            sfilesva1 = "SI"
         else
            sfext = instr(f.name,".")
            if sfext = 0 then
               sfext = "NO"
            else
               sfext = ucase(right(f.name,len(f.name)-sfext))
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
            sfext = ucase(left(f.name,sfextn))
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
              sHTML = sHTML & "<option value=" & chr(34) & f.name & chr(34)
         else
              sHTML = sHTML & "<option value=" & chr(34) & replace(newdir,"/","",1) & "/" & f.name & chr(34)
         end if

             if ccomp2 = (newdir & "/" & f.name) then   sHTML = sHTML & " selected "
              sHTML = sHTML & " >" & Ucase(newdirn & "/" & f.name) & "</option>"

      end if

  Next

  if sfcount = 0 then
              sHTML = sHTML & "<option value=" & chr(34) & ccomp & chr(34)
              sHTML = sHTML & " selected "
              sHTML = sHTML & " >No hay imagenes disponibles</option>"

   end if


  sHTML = sHTML & "</select></font>"
  sHTML = sHTML & "</FORM>" & chr(13)
'--------
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"


  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td width=50% align=center class=tt style={border-style:outset;border-width:2;}  >"
  sHTML = sHTML & "&nbsp;"
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td width=50% align=center class=tt style={border-style:outset;border-width:2;}  >"
  sHTML = sHTML & "FILTRO: "
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"


  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td width=50% align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & "&nbsp;"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "<td width=50% align=center style={border-style:outset;border-width:2;} >"
'-------------
  sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=3 style={border-style:outset;border-width:2;} ","","100%","",0,0,0,0,"","","","no")
  sHTML = sHTML & "<tr>"

  if bModificar then
      sAccion = "PRYArcSel.asp?" 
      sAccion = sAccion & "nblk=D&ord=F&Codobjeto=" & nCodigo & "&CantRegAMover=" & CantRegAMover
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&opc=" & opc & "&itemdata=" & sDocum & "&newdir=" & newdir

  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp3 ID=fcomp3 action="  & chr(34) &  sAccion  & chr(34) &  "  method=post >" & chr(13)

  sHTML = sHTML & "<td width=35% align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & ofv.GenerarInput("comp31",comp31,"agregar","","80","dt","")
  sHTML = sHTML & "</td>"


  sHTML = sHTML & "<td width=35% align=center style={border-style:outset;border-width:2;} >"

  sHTML = sHTML & "<font class=at><select  name=comp32 style={border-style:outset;border-width:4;} "
 ' sHTML = sHTML & " onchange=" & chr(34) & "fcomp3.submit();" & chr(34)

  sHTML = sHTML & " >"

  sHTML = sHTML & "<option value=" & comp32
  sHTML = sHTML & " selected " 
  sHTML = sHTML & " >" & comp32 & "</option>" 

  sHTML = sHTML & "<option value=TODOS "
  sHTML = sHTML & " >TODOS</option>" 
  sHTML = sHTML & "<option value=HTM "
  sHTML = sHTML & " >HTM</option>" 
  sHTML = sHTML & "<option value=HTML "
  sHTML = sHTML & " >HTML</option>" 
  sHTML = sHTML & "<option value=PDF "
  sHTML = sHTML & " >PDF</option>" 
  sHTML = sHTML & "<option value=MPG "
  sHTML = sHTML & " >MPG</option>" 
  sHTML = sHTML & "<option value=DOC "
  sHTML = sHTML & " >DOC</option>" 
  sHTML = sHTML & "<option value=XLS "
  sHTML = sHTML & " >XLS</option>"



  sHTML = sHTML & "</select></font>"
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td width=30% align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='FILTRAR' class=bt >"
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "</FORM>" & chr(13)

  sHTML = sHTML & "</tr>"
  sHTML = sHTML & "</table>"
'--------
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"


  sHTML = sHTML & "</table>"

  sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=3 style={border-style:outset;border-width:2;} ","","80%","",0,0,0,0,"","","","no")
  sHTML = sHTML & "<tr>"
   opc = "NO"
if opc = "NO" then
  sHTML = sHTML & "<td>"
  sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=3 style={border-style:outset;border-width:2;} ","","100%","",0,0,0,0,"","","","no")



  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  class=tt align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & "Documento&nbsp;anterior"
  sHTML = sHTML & "</td>"



  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td   class=cc   >"



     ncomp =  "PROYECTOS/" &  session("NCliente") &  "/" & ccomp

  sHTML = sHTML & "<font color=#000080 ><B>" & ncomp & "</b></font>"
  sHTML = sHTML & "</td>"


  sHTML = sHTML & "</tr>"

'  sHTML = sHTML & "<tr>"
'  sImage = ofv.GenImage("Imagen","aaa width=100 border=1 ","","", "legajos/" & session("NCliente") & "/" & ccomp,"","")
'  sHTML = sHTML & ofv.GenCelda("aaa ","","center","","","","",sImage,"si")
'  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "</table>"
  sHTML = sHTML & "</td>"

  xcolsp = ""

else
  xcolsp = " colspan=2 "

end if
  sHTML = sHTML & "<td " & xcolsp & " >"
  sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=3 style={border-style:outset;border-width:2;} ","","100%","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  class=tt align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & "Documento&nbsp;nuevo"
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "</tr>"

'  if nblk = "D" then ccomp2 = ccomp
'  if nblk = "S" then ccomp2 = ccompS
'  if nblk = "T" then ccomp2 = ccompT

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  class=cc   >"


     ncomp =  "PROYECTOS/" & session("NCliente") & "/" & ccomp2

  sHTML = sHTML & "<font color=#993333 ><B>" & ncomp & "</b></font>"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

'  sHTML = sHTML & "<tr>"
'  sImage = ofv.GenImage("Imagen","aaa width=100 border=1 ","","","legajos/" & session("NCliente") & "/" & ccomp2,"","")
'  sHTML = sHTML & ofv.GenCelda("aaa","","center","","","","",sImage,"si")
'  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "</table>"
  sHTML = sHTML & "</td>"

'''''''''''''''''''''

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td align=center  >"
  if bModificar then

      sAccion = "asp/Actualizar.asp?fname=" & sDocum & "&Tabla=PROYantec&volver=../PRYARCHIVOS.asp?"
      sAccion = sAccion & "Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo



  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=acep ID=acep action=" & sAccion & "  method=post >" & chr(13)
  sHTML = sHTML & ofv.GenerarInput(rs(3).name,ccomp2,"hidden","","10","","")
  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='ACEPTAR' class=bt >"
  sHTML = sHTML & "</FORM>" & chr(13)
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td  align=center   >"

    sAccion = "PRYARCHIVOS.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sHTML = sHTML & "<FORM  name=Cancel ID=Cancel action=" & sAccion & "  method=post >" & chr(13)
  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='CANCELAR' class=bt >"
  sHTML = sHTML & "</FORM>" & chr(13)
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "</tr>"
  sHTML = sHTML & "</table>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</center></body></html>"
end if

response.write sHTML

%>