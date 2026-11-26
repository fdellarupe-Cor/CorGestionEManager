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
  response.write bModificar
  if session("Nomodificar") = "S" then bModificar = False

  nblk = request.querystring("nblk")
  if len(nblk) = 0 then nblk = "S"

  nAccion = request.querystring("idtxt") &""
  
  if len(naccion)=0 then nAccion=request.querystring("Accion")
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  vueltaC = request.querystring("volver")

  if session("volverproc") = "" then session("volverproc") = vueltaC
  if vueltaC = "" then vueltaC = session("volverproc")

 ' volver2 = ofv.vueltaasp2("../") & "_Accion=" & nAccion
  volver2 = ofv.vueltaaspEx("../","0","0") & "_Accion=" & nAccion

  esttxt = " STYLE={font-family:tahoma;color:#4466bb;font-size:10pt;font-weight:bold;background-color:#ffffff;} "

  set cn = ofv.conectar(ofv.strconn2)
  Strsql = "SELECT * From objestdet where objid = " & nAccion
  response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     wproced = nAccion 'rs(3)
     waccion = nAccion
     wlink = rs("detimagen") &""
     if isnull(wlink) or len(wlink)=0 then wlink = "0"
     'wtxtl = rs(20)
     response.write wlink & "   " & nblk
  end if
  wlink=clng(wlink)
  call ofv.cerrarconsulta(rs)

    sHTML = ofv.FormBody & "<center><br>"


' inserta link local

  if nblk = "IL" then

     wlink = request.form("comp1") 

     Strsql = "SELECT * From objetos where codobjeto = " & wlink
     'RESPONSE.WRITE "<BR>" & STRSQL & "<BR>"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
        wtipoL = rs(2)
     else
        wtipoL = "NO"
        wlink = 0 
     end if
     call ofv.cerrarconsulta(rs)

     if wlink > 0 then
       
        strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2, codobjetoL, tipobjetoL, Lurl, LurlDesc, LTipo "
        strsql = strsql & ") Values (" & nCodigo & ",'" & sTipo & "',"
        strsql = strsql & "'" & wproced & "','" & waccion & "',"
        strsql = strsql & wlink & ",'" & wtipoL & "','','','L'"
        strsql = strsql & ")"
        RESPONSE.WRITE "<BR>" & STRSQL & "<BR>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        Strsql = "SELECT * From OBJLNK where iden1 = '" & wproced & "' and iden2 = '" & waccion & "' "  
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rs.eof then
           wlnkid = rs(0)
        else
           wlnkid = 0 
           wlink = 0 
        end if
        call ofv.cerrarconsulta(rs)

        if wlnkid > 0 then
           strsql = "update objestdet set detimagen = " & wlnkid
           strsql = strsql & " where  objid = " & nAccion
           response.write strsql & "<br>" & wproced & " 222"
           call ofv.crearconsultaEx(StrSql,cn,1,volver)
        end if
     end if

     nblk = "S"
  end if 



' inserta link internet

  if nblk = "II" then

     wlink = request.form("comp1") 


        strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2, codobjetoL, tipobjetoL, Lurl, LurlDesc, LTipo "
        strsql = strsql & ") Values (" & nCodigo & ",'" & sTipo & "',"
        strsql = strsql & "'" & wproced & "','" & waccion & "',"
        strsql = strsql & "0,'','" & wlink & "','" & wlink & "','I'"
        strsql = strsql & ")"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        Strsql = "SELECT * From OBJLNK where iden1 = '" & wproced & "' and iden2 = '" & waccion & "' "  
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rs.eof then
           wlnkid = rs(0)
        else
           wlnkid = 0 
           wlink = 0 
        end if
        call ofv.cerrarconsulta(rs)

        if wlnkid > 0 then
           strsql = "update objestdet set detimagen = " & wlnkid
           strsql = strsql & " where  objid = " & nAccion
           call ofv.crearconsultaEx(StrSql,cn,1,volver)
        end if

     wlink = wlnkid
     nblk = "S"
  end if 




' agrega link local

  if nblk = "AL" then

  sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3 style={border-style:outset;border-width:2;} ","","80%","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=2 class=tt align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & "Agregar&nbsp;Link&nbsp;a&nbsp;documento "
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"
  sAccion =  replace(vueltaC,"../","")
  sAccion =  replace(sAccion,"_","&")
  sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "</FORM>" & chr(13)

'-------------
  if bModificar then
      sAccion = "EdProcedLnk2.asp?" 
      sAccion = sAccion & "Accion=" & nAccion  & "&nblk=IL&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)


  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=2 align=center style={border-style:outset;border-width:2;} >"

  sHTML = sHTML & "<font class=at><select size=10  name=comp1 style={border-style:outset;border-width:4;} "
  sHTML = sHTML & " >"

      Strsql = "Select * From csGenObjetos where estado = 3  OR ESTADO = 1 order by xdescri"
 '     sInput = ofv.GenerarCombo("document.edGeneral",rs(4).name,rs(4).value,Strsql,cn,0,1,"sb")

'  Strsql = "SELECT tipoaccionalt,descri From TIPOACCALT  where modo = '" & ccomp3 & "' Order By descri"
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  do while not rs2.eof
     sHTML = sHTML & "<option value=" & rs2(0)
     sHTML = sHTML & " > " & rs2(1) & "</option>"

     rs2.movenext
  loop
  call ofv.cerrarconsulta(rs2)
  sHTML = sHTML & "</select></font>"

'--------
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  align=center   >"

  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='Aceptar' class=bt >"

  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td  align=center   >"
  sHTML = sHTML & "<input  type=Button  size=10%  value='Cancelar' class=bt "
  sHTML = sHTML & " onclick=" & chr(34) & "fcomp2.submit();" & chr(34) &  " >"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "</FORM>" & chr(13)
  sHTML = sHTML & "</table><br><br>"


     nblk = "NO"

  end if 



' agrega link internet

  if nblk = "AI" then

  sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3 style={border-style:outset;border-width:2;} ","","80%","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=2 class=tt align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & "Agregar&nbsp;Link&nbsp;a&nbsp;Internet "
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"
  sAccion =  replace(vueltaC,"../","")
  sAccion =  replace(sAccion,"_","&")
  sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "</FORM>" & chr(13)

'-------------
  if bModificar then
      sAccion = "EdProcedLnk2.asp?" 
      sAccion = sAccion & "Accion=" & nAccion  & "&nblk=II&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)


  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=2 align=center style={border-style:outset;border-width:2;} >"

  sHTML = sHTML & ofv.GenerarInput("comp1","www.","agregar","","100","dt","sb")

'--------
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  align=center   >"

  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='Aceptar' class=bt >"

  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td  align=center   >"
  sHTML = sHTML & "<input  type=Button  size=10%  value='Cancelar' class=bt "
  sHTML = sHTML & " onclick=" & chr(34) & "fcomp2.submit();" & chr(34) &  " >"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "</FORM>" & chr(13)
  sHTML = sHTML & "</table><br><br>"


     nblk = "NO"

  end if 


' Alta 1era vez


  if wlink = 0 and nblk = "S" then

  sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=5 style={border-style:outset;border-width:2;} ","","80%","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=3 class=tt align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & "Agregar&nbsp;Link "
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

'-------------

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  COLSPAN=3 align=center   >&nbsp;</TD>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  align=center   >"

  if bModificar then
      sAccion = "EdProcedLnk2.asp?" 
      sAccion = sAccion & "Accion=" & nAccion  & "&nblk=AL&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)

  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='AGREGAR LINK A OTRO DOCUMENTO' class=bt >"

  sHTML = sHTML & "</FORM>" & chr(13)

  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td  align=center   >"

  if bModificar then
      sAccion = "EdProcedLnk2.asp?" 
      sAccion = sAccion & "Accion=" & nAccion  & "&nblk=AI&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp3 ID=fcomp3 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)

  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='AGREGAR LINK A INTERNET' class=bt >"

  sHTML = sHTML & "</FORM>" & chr(13)

  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td  align=center   >"
  sAccion =  replace(vueltaC,"../","")
  sAccion =  replace(sAccion,"_","&")
  sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "<input  type=submit  size=10%  value='CANCELAR' class=bt >"
  sHTML = sHTML & "</FORM>" & chr(13)
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  COLSPAN=3 align=center   >&nbsp;</TD>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "</table><br><br>"


     nblk = "NO"


  end if





' modicar link local


  if nblk = "ML" then

     wlink = request.form("comp1") 

     Strsql = "SELECT * From objetos where codobjeto = " & wlink
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
        wtipoL = rs(2)
     else
        wtipoL = "NO"
        wlink = 0 
     end if
     call ofv.cerrarconsulta(rs)

     if wlink > 0 then

        strsql = "update OBJLNK set codobjetoL = " & wlink & ", tipobjetoL = '" & wtipoL & "' "
        strsql = strsql & "where iden1 = '" & wproced & "' and iden2 = '" & waccion & "' "
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

     end if

     nblk = "S"
  end if 

' modicar link internet


  if nblk = "MI" then

     wlink = request.form("comp1") 

     wlinkD = request.form("comp2") 



        strsql = "update OBJLNK set Lurl = '" & wlink & "', LurlDesc = '" & wlinkD & "' "
        strsql = strsql & "where iden1 = '" & wproced & "' and iden2 = '" & waccion & "' "
        call ofv.crearconsultaEx(StrSql,cn,1,volver)



     nblk = "S"
  end if 



' anular link


  if nblk = "A" then

     response.write "anular link: " & "<br>" 

     sAccion =  replace(vueltaC,"../","")
     sAccion =  replace(sAccion,"_","&") 
     'sAccion =  replace(sAccion,"***","&")
     
    sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
     

     strsql = "delete from OBJLNK "
     strsql = strsql & "where iden1 = '" & wproced & "' and iden2 = '" & waccion & "' "
     call ofv.crearconsultaEx(StrSql,cn,1,volver)

 

    ' strsql = "update OBJaccion set rellink = 0, reflink = '' "
    
      strsql = "update objestdet set detimagen =NULL" 
      strsql = strsql & " where  objid = " & nAccion
      'strsql = strsql & " where  objid = " & nAccion
     call ofv.crearconsultaEx(StrSql,cn,1,volver)
      'RESPONSE.WRITE SACCION
     response.redirect sAccion

  '   nblk = "S"
     nblk = "NO"
  end if 




' Mostrar link


  if nblk = "S" then
 
  'Strsql = "SELECT * From objaccion where objid = " & nAccion
  
  Strsql = "select *  From objestdet where objid=" &  nAccion
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     'wproced = rs(3)
     'waccion = rs(18)
    ' wlink = rs(19)
     wlink=rs("DETIMAGEN") &""
     if isnull(wlink) or len(wlink)=0 then wlink = "0" 
     'wtxtl = rs(20)
     wlink=clng(wlink)

     wtipoL = 0
     wUrl = "NO"
     wUrlDesc = "NO"
     wtipo = "NO"

     Strsql = "SELECT * From objlnk where objid = " & wlink
     set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs1.eof then
        wtipoL = rs1(6)
        wUrl = rs1(8)
        wUrlDesc = rs1(9)
        wtipo = rs1(10)

     end if
     call ofv.cerrarconsulta(rs1)

 ' response.write  wtipoL &  wUrl & wUrlDesc & wtipo & "<br>"


  sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3 style={border-style:outset;border-width:2;} ","","80%","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=2 class=tt align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & "Link&nbsp;a&nbsp;documento"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sAccion =  replace(vueltaC,"../","")
  sAccion =  replace(sAccion,"_","&")
  sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "</FORM>" & chr(13)

  if bModificar then
      sAccion = "EdProcedLnk2.asp?" 
      sAccion = sAccion & "Accion=" & nAccion  & "&nblk=A&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp3 ID=fcomp3 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "</FORM>" & chr(13)

  if wtipo = "I" then

' muestra link Internet

'-------------
  if bModificar then
      sAccion = "EdProcedLnk2.asp?" 
      sAccion = sAccion & "Accion=" & nAccion  & "&nblk=MI&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=2 class=cc align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & ofv.GenerarInput("comp1",WUrl,"textsb","fcomp1","100","dt","sb")
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=2 class=tt align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & "Descripcion&nbsp;para&nbsp;el&nbsp;link"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"  & chr(13)

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=2 class=cc align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & ofv.GenerarInput("comp2",wUrlDesc,"textsb","fcomp1","100","dt","sb")
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "</FORM>" & chr(13)

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=2 class=tt align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & "Referencia&nbsp;para&nbsp;el&nbsp;link"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"  & chr(13)

  if bModificar then
     sAccion = "asp/ActualizarEN.asp?ID=" & nAccion & "&Tabla=Objaccion&volver=" & volver2
     sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp0 ID=fcomp0 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=2 class=cc align=center style={border-style:outset;border-width:2;} >"

    t3b = "<font class=at><textarea "
    t3b = t3b & "name=" & rs(8).name & " cols=90 rows=6 " & esttxt
    if bModificar then
      t3b = t3b & " onblur=fcomp0.submit() "
      t3b = t3b & " onkeypress=" & chr(34)
      t3b = t3b & " if (this.value.length > 490) {alert('El texto es demasiado"
      t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
    else
      t3b = t3b & " readonly >"
    end if
    xdatos = wtxtl
    t3b = t3b & xdatos & "</textarea></font>"


  sHTML = sHTML & t3b
  sHTML = sHTML & ofv.GenerarInput(rs(6).name,rs(6),"hidden","","10","","")
  sHTML = sHTML & ofv.GenerarInput(rs(4).name,rs(4),"hidden","","10","","")
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

'  sHTML = sHTML & "</FORM>" & chr(13)


  else


' muestra link local

  Strsql = "Select * From csGenObjetos where codobjeto = " & wtipoL
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs2.eof then
     wNomL = rs2(1)
  else
     wNomL = "S/D"
  end if
  call ofv.cerrarconsulta(rs2)  

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=2 class=cc align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & "<b>" & wNomL & "</b>"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=2 class=tt align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & "Referencia&nbsp;para&nbsp;el&nbsp;link"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"  & chr(13)

  if bModificar then
     sAccion = "asp/ActualizarEN.asp?ID=" & nAccion & "&Tabla=Objaccion&volver=" & volver2
     sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp0 ID=fcomp0 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=2 class=cc align=center style={border-style:outset;border-width:2;} >"

    t3b = "<font class=at><textarea "
    't3b = t3b & "name=" & rs(20).name & " cols=90 rows=6 " & esttxt
    t3b = t3b & "name=" & rs(8).name & " cols=90 rows=6 " & esttxt
    if bModificar then
      t3b = t3b & " onblur=fcomp0.submit() "
      t3b = t3b & " onkeypress=" & chr(34)
      t3b = t3b & " if (this.value.length > 490) {alert('El texto es demasiado"
      t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
    else
      t3b = t3b & " readonly >"
    end if
    xdatos = wtxtl
    t3b = t3b & xdatos & "</textarea></font>"


  sHTML = sHTML & t3b
  sHTML = sHTML & ofv.GenerarInput(rs(6).name,rs(6),"hidden","","10","","")
  sHTML = sHTML & ofv.GenerarInput(rs(4).name,rs(4),"hidden","","10","","")
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "</FORM>" & chr(13)

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=2 class=tt align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & "Modificar&nbsp;el&nbsp;link"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"  & chr(13)

'-------------
  if bModificar then
      sAccion = "EdProcedLnk2.asp?" 
      sAccion = sAccion & "Accion=" & nAccion  & "&nblk=ML&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)


  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=2 align=center style={border-style:outset;border-width:2;} >"

  sHTML = sHTML & "<font class=at><select size=10  name=comp1 style={border-style:outset;border-width:4;} "
  sHTML = sHTML & " onchange=" & chr(34) & "fcomp1.submit();" & chr(34)
  sHTML = sHTML & " >"

  Strsql = "Select * From csGenObjetos where estado = 3  OR ESTADO = 1 order by xdescri"
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  do while not rs2.eof
     sHTML = sHTML & "<option value=" & rs2(0)
     if rs2(0) = wtipoL then sHTML = sHTML & " selected "
     sHTML = sHTML & " > " & rs2(1) & "</option>"

     rs2.movenext
  loop
  call ofv.cerrarconsulta(rs2)
  sHTML = sHTML & "</select></font>"

'--------
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"


  end if


  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  align=center   >"

  sHTML = sHTML & "<input  type=Button size=10%  value='Anular Link' class=bt "
  sHTML = sHTML & " onclick=" & chr(34) & "fcomp3.submit();" & chr(34) &  " >"
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td  align=center   >"
  sHTML = sHTML & "<input  type=Button  size=10%  value='Volver' class=bt "
  sHTML = sHTML & " onclick=" & chr(34) & "fcomp2.submit();" & chr(34) &  " >"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "</FORM>" & chr(13)
 
  sHTML = sHTML & "</table><br><br>"

  end if

  call ofv.cerrarconsulta(rs)

  end if 



  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</center></body></html>"
end if
  if nblk <> "A" then response.write sHTML

%>