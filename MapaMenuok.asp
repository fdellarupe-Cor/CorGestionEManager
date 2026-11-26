<!-- #INCLUDE FILE="coress20/IncFile/CorConect.inc" -->

<%
xx=ARMAM()

response.write xx
Function ARMAM()
	Dim cbmm

    dim str
    DIM AUsrG
	str="" & Chr(13)
	 'str=str & chr(13) & "<li class=lixstk TITLE='' ><a  Xclass=lixstk  href=# class='link'  > <i style=' font-size: 25px; ' class='la la-user'></i> <font size=2 color=  #fdfefe   face=verdana>USUARIO</font></a>" & Chr(13) & "</li>" & Chr(13)
     XMESAURL = ""

 If session("URLTYP") = "V" Then
	XMESAURL = Left(request.servervariables("SCRIPT_NAME"), InStr(Right(request.servervariables("SCRIPT_NAME"), Len(request.servervariables("SCRIPT_NAME")) - 1), "/"))
 End If


  If session("URLPORT") = "N" Then
	 XMESAURL = tiposite("") & request.servervariables("SERVER_NAME") & XMESAURL
  Else
	 XMESAURL = tiposite("") & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL
  End If



	surx=XMESAURL & "/coremanager/novedmenu.asp"

    suserinf="<i style=' font-size: 35px; ' class='la la-user'></i>  <font style=' font-size: 20px;'    face=calibri><i>" &  capitalize(Session("SegUsuario")) & "</i></font><br>"

    xx=getinfouser(UserLogon())
    suserinf=suserinf & "<i style=' font-size: 18px;'>Email: " & capitalize(UCase(Session("sgu_mail")& "" ))
    suserinf=suserinf & "<BR>Grupos: "

    wktgrupos = Split(session("SegGrpdsc"), Chr(9))

	For wkti = 0 To UBound(wktgrupos)
		suserinf=suserinf & " " & capitalize(wktgrupos(wkti))
	NEXT

    suserinf=suserinf & "</i>"
    suserinf=""

    ARMAM="<div id=contm><ul>" & chr(13)
   ' ARMAM=ARMAM & "<li  TITLE='' ><a class=userc  href=#   xstyle='background-color:#FF9B42' >  <font size=2 color=  #fdfefe    face=calibri>" & Session("SegUsuario")&"&nbsp;&nbsp;<i style=' font-size: 25px; ' class='la la-user'></i></font></a>" & Chr(13) & "</li>" & Chr(13)

	if  Session("DELLOG") = "S" then
		'  session("SegGrpLst")="MAS" & CHR(9) & "XX"
		   AUsrG = Split(session("SegGrpLst"),Chr(9))
		   AUsrG2 = Split(session("SegGrpLst"),Chr(9))
		   XMFSEP=""
		   For XMFI = 0 To UBound(AUsrG)
		   		UsrGS  = UsrGS  & XMFSEP & "'" & Replace(AUsrG(XMFI) & "","'","''") & "'"
		   		XMFSEP  = ","
		   Next
		  else
		  Session("seggrp")="MAS" & CHR(9)
		   UsrGS=  "'" & Session("seggrp")  & "'"
		  ' AUsrG2 = Split(Session("seggrp") & chr(9) & "22222222222222222" ,Chr(9))
		   AUsrG2 = Split(Session("seggrp") ,Chr(9))
		   'response.write(AUsrG2(0))
	 end if

	Set cmm = OFV.conectar(OFV.strconn0)
    sub1=""
   ' wkaplic="EM"
     SQL= "Select max (BB.codapli) as codapli ,MAX(isnull(  CAST(BB.atrGrupo AS INT)  ,1)) as atrGrupo ,MAX(isnull(CAST(BB.atrUsr AS INT) ,1) )as atrUsr "
     SQL=SQL & " from ( "
     XMFSEP=""
     SQL2=""
     For XMFI = 0 To UBound(AUsrG2)
     UsrGS  ="'" & Replace(AUsrG2(XMFI) & "","'","''") & "'"
     sql2=sql2 & chr(13) & XMFSEP & chr(13)
     sql2=sql2 & "Select b.codapli as codapli ,isnull(  CAST(c.grpestado AS INT)  ,1) as atrGrupo ,isnull(CAST(d.usrEstado AS INT) ,1) as atrUsr"
	 sql2=sql2 &  " FROM CorAplic b "
	 sql2=sql2 &  " left join atrmXgrupo c on   ( c.Tipo = 'A' and c.grupo IN("& UsrGS &") and  b.codapli=C.APLICACION) "
	 sql2=sql2 &  " left join atrmXusuario d on   ( D.Tipo = 'A' and d.usuario='" & Session("segperfil") & "' and D.APLICACION=b.codapli) "
     sql2=sql2 &  " where b.enable <> 0"
     XMFSEP= CHR(13) & " UNION " & CHR(13)
     NEXT
     SQL =SQL & SQL2 & ") BB"
     SQL=SQL & " GROUP BY BB.codapli " & CHR(13) & CHR(13)

   ' sql=" Select max (b.codapli) as codapli ,MAX(isnull(  CAST(c.grpestado AS INT)  ,1)) as atrGrupo ,MAX(isnull(CAST(d.usrEstado AS INT) ,1) )as atrUsr"
   ' sql=  SQL & " FROM CorAplic b "
   ' sql=  SQL & " left join atrmXgrupo c on   ( c.Tipo = 'A' and c.grupo IN(" & UsrGS &") and  b.codapli=C.APLICACION) "
   ' sql=  SQL & " left join atrmXusuario d on   ( D.Tipo = 'A' and d.usuario='" & Session("segperfil") & "' and D.APLICACION=b.codapli) "
   ' sql=  SQL & " where b.enable <> 0 "
   ' sql=  SQL & " GROUP BY b.codapli"
  ' RESPONSE.WRITE SQL
    Set rsApp = ofv.crearconsultaEx(sql,cmm,1,parametros)
    AppH=""
    XMFSEP=""
    do while not rsapp.eof
       if rsAPP("atrGrupo") <> 0 and rsAPP("atrUsr") <> 0 then
        AppH  = AppH  & XMFSEP & "'" & rsapp(0) & "'"
        XMFSEP  = ","
       END IF

      rsapp.movenext
    loop
    call ofv.cerrarconsulta(rsapp)

	ARMAM= ARMAM &  listarM2("0",0,cmm,AppH,UsrGS)

	ARMAM= ARMAM & "</nav></div>" & sub1
	If Session("UsrCorp") and Session("EMPMUN") = "S" Then
	 ARMAM= ARMAM  & chr(13) & "<div class=nnn title='Empresas' id=submEmp data-width='410' data-height='330' style=""display:nonex;height:280px;width:380px"">" & selecEmpre() & "</DIV>"
	end if
	ARMAM=ARMAM & "</div>"

End Function

function  ListarM2(mdepen,nivel,cmm,aplic,UsrgS)
	Dim rs,rs1
	Dim rsoef



	'	sql="Select Menux.*,titapli From Menux, CorAplic where CorAplic.codapli = Menux.mcodapli and CorAplic.CodApli in( " & aplic & " ) AND Menux.mdepen = '" & mdepen & "'"
	'	sql= sql & "ORDER BY CorAplic.CodApli"





	'	sql="Select max(a.id), a.MENU,a.mcodapli,b.titapli,max(a.menuatr) as menuatr,MAX(isnull(  CAST(c.grpestado AS INT)  ,1)) as atrGrupo ,MAX(isnull(CAST(d.usrEstado AS INT) ,1) )as atrUsr, max(a.morden) as morden,max(ordenapli )  as ordenapli,max(asp )  as asp"
	'	sql= sql  &  " From Menux a "
	'	sql=sql & "inner join CorAplic b on ( b.codapli = a.mcodapli ) "
	'	sql=sql & " left join atrmXgrupo c on   (c.alias=a.menuatr and c.Tipo = 'S' and c.grupo IN(" & UsrGS &") and c.aplicacion=b.codapli) "
	'	sql =sql & " left join atrmXusuario d on   (d.alias=a.menuatr and D.Tipo = 'S' and d.usuario='" & Session("segperfil") &"' and D.aplicacion=b.codapli) "
  '      sql= sql & " where  b.CodApli in( " & aplic & " ) AND a.mdepen = '0'"
'
  '      ' no trae formularios
  '      if session("CORMODUSOLABM") <> "S" then
  '       sql= sql & " and menuatr not in('CLIS245159671557361072846976','CLIS45312740033826860617908')"
  '      end if
  '      sql= sql & " GROUP BY b.titapli,a.mcodapli,a.MENU"
  '      sql= sql & " order by  ordenapli,b.titapli,a.mcodapli,a.MENU,morden"


    sql="Select max(bb.id), bb.MENU,bb.mcodapli,bb.titapli,max(bb.menuatr) as menuatr,MAX(isnull(  CAST(bb.atrGrupo AS INT)  ,1)) as atrGrupo ,MAX(isnull(CAST(bb.atrUsr AS INT) ,1) )as atrUsr, max(bb.morden) as morden,max(bb.ordenapli )  as ordenapli,max(bb.asp )  as asp"
    sql =sql & " from ("

     XMFSEP=""
     SQL2=""
        For XMFI = 0 To UBound(AUsrG2)
         UsrGS  ="'" & Replace(AUsrG2(XMFI) & "","'","''") & "'"
         sql2=sql2 & chr(13) & XMFSEP & chr(13)
         sql2=sql2 &  "Select a.id, a.MENU,a.mcodapli,b.titapli,a.menuatr as menuatr,isnull(  CAST(c.grpestado AS INT)  ,1) as atrGrupo ,isnull(CAST(d.usrEstado AS INT) ,1) as atrUsr, a.morden as morden,ordenapli   as ordenapli,asp  as asp "

		 sql2=sql2 & "From Menux a "
		 sql2=sql2 & " inner join CorAplic b on ( b.codapli = a.mcodapli ) "
		 sql2=sql2 & " left join atrmXgrupo c on   (c.alias=a.menuatr and c.Tipo = 'S' and c.grupo IN("& UsrGS & ") and c.aplicacion=b.codapli) "
		 sql2=sql2 & " left join atrmXusuario d on   (d.alias=a.menuatr and D.Tipo = 'S' and d.usuario='" & Session("segperfil") & "' and D.aplicacion=b.codapli) "
		 sql2=sql2 & " where  b.CodApli in(" & aplic & " ) AND a.mdepen = '0'"
		    ' no trae formularios
		 if session("CORMODUSOLABM") <> "S" then
          sql2=sql2 & " and menuatr not in('CLIS245159671557361072846976','CLIS45312740033826860617908')"
         end if
         XMFSEP= CHR(13) & " UNION " & CHR(13)
        NEXT
     SQL =SQL & SQL2 & ") BB"
    sql =sql & "  GROUP BY bb.titapli,bb.mcodapli,bb.MENU"
    sql =sql & "  order by  ordenapli,bb.titapli,bb.mcodapli,bb.MENU,morden"


	'   response.write sql & "<br>"

	Set rs = ofv.crearconsultaEx(sql,cmm,1,parametros)

	If  rs.eof Then
		rseof="S"
	Else
		rseof="N"
	End If

	Dim a(100,1) , i
	Dim imenu , iSm
	' Recorre los subdirectorios
	i=0
	 If Session("UsrCorp") and Session("EMPMUN") = "S" Then
       str=str & chr(13) & "<li class=lixstk TITLE='Empresas' ><a  Xclass=lixstk  xtitle='Empresas' href=# class='link'  data-id='submEmp' > <i style=' font-size: 25px; ' xclass='la la-pencil'></i> <font size=2 color=  #fdfefe   face=verdana><I>EMPRESAS</I></font></a>" & Chr(13) & "</li>" & Chr(13)
     end if
	Do While Not rs.eof
		'  ReDim Preserve a(i,0)

        if rs("atrGrupo") <> 0 and rs("atrUsr") <> 0 then
'if true then
		sql2="Select top 1  Menux.* From Menux, CorAplic where CorAplic.codapli = Menux.mcodapli and CorAplic.CodApli in( " & aplic & " ) AND Menux.mdepen = '" & rs("menuatr") & "'"
		Set rs2 = ofv.crearconsultaEx(sql2,cmm,1,parametros)

		If Not rs2.eof Then
     '   response.write sql2 & chr(13)
        if rs(2)="COREM20" THEN
          str=str & chr(13) & "<li xclass=lixstk TITLE='" & rs("titapli") &"' ><a  Xclass=lixstk  xtitle='" & rs("titapli") & "' href=# class='link'  data-id='subm" & rs(0) & "' > <i xstyle=' font-size: 25px; ' xclass='la la-pencil'></i> <font size=2 xcolor=  #fdfefe   xxface=calibri>" & UCase(rs("menu")) & "</font></a>" & Chr(13) & "</li>" & Chr(13)
        ELSE
		  str=str & chr(13) & "<li xclass=lixstk TITLE='" & rs("titapli") &"' ><a  Xclass=lixstk  xtitle='" & rs("titapli") & "' href=# class='link'  data-id='subm" & rs(0) & "' > <i xstyle=' font-size: 25px; ' xclass='la la-gear'></i> <font size=2 xcolor=  #fdfefe   xface=calibri>" & UCase(rs("menu")) & "</font></a>" & Chr(13) & "</li>" & Chr(13)
		END IF

       'sub menu
        sub1=sub1 & armasubm(rs("menuatr"),cmm,rs(1),1,"'" & rs(2) & "'" ,rs(0),UsrGS)

		Else
			xsel="menuatr"
			sInput = Trim(rs(xsel))
			sHref = Trim(rs("asp"))
			sTarget = "info"
			sPath = Session("CorPath")
			If rs(3) <> "" Then sTarget = rs(3)
			spt = InStr(rs("asp"), ".asp")
			xesm = "N"
			If spt > 1 Then
				If Left(rs("asp"), spt - 1) = "Submenus" Then
					sPath = "../coress20/"
					xesm = "S"
				End If
			End If
			sseg = InStr(sHref, "?")
			If sseg = 0 Then
				sHref = sHref & "?"
			Else
				sHref = sHref & "&"
			End If
			sHref = sHref & "seg=" & OFV.convertircar(sInput, " ", "_") & "&MNN=" & OFV.convertircar(UCase(rs(1)), " ", "_")
			ssMnn = "MNN" & OFV.convertircar(UCase(rs(1)), " ", "_")
			str=str & chr(13)  & "<LI class=lixstk><font color=  #fdfefe   size=1 face=calibri ><a onclick=" & Chr(34) & "SetMnn('" & ssMnn & "');" & Chr(34)  & " class=xxxx target=info href='" & sHref & "'>" & UCase(RS("MENU")) & "</a></font></LI>"  & Chr(13)
		End If
        end if


		rs.movenext

	Loop
	 call ofv.cerrarconsulta(rs)
	str=str & "</UL>" & Chr(13)
	ListarM2=str
End function



function armasubm(id,cmm,idx,nivel,aplic,idx2,UsrGS)
dim ss7
'dim AUsrG
'dim rsx
'dim XMFI

'sql="Select Menux.*,dirapli,titapli From Menux, CorAplic where CorAplic.codapli = Menux.mcodapli and CorAplic.CodApli in( " & aplic & " ) AND Menux.mdepen = '" & id & "' "
SQL = "Select max(BB.id) as id, BB.MENU,BB.mcodapli,BB.titapli,max(BB.menuatr) as menuatr,MAX(isnull(  CAST(BB.atrGrupo AS INT)  ,1)) as atrGrupo ,MAX(isnull(CAST(BB.atrUsr AS INT) ,1) )as atrUsr, max(BB.morden) as morden,max(BB.ordenapli )  as ordenapli,max(BB.asp )  as asp"
SQL= SQL & ",max(BB.target) as target,max(BB.dirapli) as dirapli"
sql=sql  & " FROM ( " & chr(13)

'   if  Session("DELLOG") = "S" then
'    AUsrG = Split(session("SegGrpLst"),Chr(9))
'   else
'   Session("seggrp")=Session("seggrp") & chr(9) & Session("seggrp") & "222"
'    AUsrG = Split(Session("seggrp"),Chr(9))
'   end if

   XMFSEP=""
   sql2=""
   For XMFI = 0 To UBound(AUsrG2)
    UsrGS  =  "'" & Replace(AUsrG2(XMFI) & "","'","''") & "'"

    sql2=sql2 & chr(13) & XMFSEP & chr(13)
	sql2=sql2 & "	Select a.id as id, a.MENU,a.mcodapli,b.titapli,a.menuatr as menuatr,isnull(  c.grpestado   ,1) as atrGrupo ,isnull(CAST(d.usrEstado AS INT) ,1)as atrUsr, a.morden as morden,ordenapli   as ordenapli,asp  as asp"
	sql2=sql2 & "	, Dirapli  ,target as target"
	sql2=sql2 & "	 From Menux a "
	sql2=sql2 & "	inner join CorAplic b on ( b.codapli = a.mcodapli ) "
	sql2=sql2 & "	left join atrmXgrupo c on   (c.alias=a.menuatr  and c.grupo IN("& UsrGS &") and c.aplicacion=b.codapli) "
	sql2=sql2 & "	left join atrmXusuario d on (d.alias=a.menuatr  and d.usuario='"& Session("segperfil") &"' and D.aplicacion=b.codapli)"
	sql2=sql2 & "     where  b.CodApli in( " & aplic &" ) AND a.mdepen = '" & id& "'"
	'sql2= chr(13) &  XMFSEP & chr(13) & sql2
	 XMFSEP  = "UNION"
   Next
sql= sql & sql2 & chr(13) & " )BB " & chr(13)
sql=sql & " GROUP BY BB.titapli,BB.mcodapli,BB.MENU"
sql=sql & " order by  morden" & chr(13)  & chr(13)

'sql= sql  &  " From Menux a "
'sql=sql & "inner join CorAplic b on ( b.codapli = a.mcodapli ) "
'sql=sql & " left join atrmXgrupo c on   (c.alias=a.menuatr  and c.grupo IN(" & UsrGS &") and c.aplicacion=b.codapli) "
'sql =sql & " left join atrmXusuario d on   (d.alias=a.menuatr  and d.usuario='" & Session("segperfil") &"' and D.aplicacion=b.codapli) "

'sql= sql & " where  b.CodApli in( " & aplic & " ) AND a.mdepen = '" & id & "'"
'sql= sql & " GROUP BY b.titapli,a.mcodapli,a.MENU"
'sql= sql & " order by  morden"
ss7=""
'if ID="CONFIGURACION" THEN
'response.write sql & "<BR><BR>"
'END IF

 Set rsx = ofv.crearconsultaEx(sql,cmm,1,parametros)

 if not rsx.eof then
 if nivel=1 then
  ss7=ss7 & chr(13) & "<div class=nnn title='" & UCase(idx) & "'" & "id=subm"  & idx2  & " style='display:nonex;'><ul >"
 else
  ss7=ss7 & chr(13) & "<div class=nnn title='" & UCase(idx) & "'" & "id=subm"  & idx2  & " style='xdisplay:nonex;'><ul class=ulx>"
 end if

 Do While Not rsx.eof

  ' response.write "222"
   if rsx("atrGrupo") <> 0 and rsx("atrUsr") <> 0 then
    STEXTO1=UCase(rsx("menu"))
    xsel="menuatr"
	sInput = Trim(rsx(xsel))
	     XMESAURL = ""

	 If session("URLTYP") = "V" Then
		XMESAURL = Left(request.servervariables("SCRIPT_NAME"), InStr(Right(request.servervariables("SCRIPT_NAME"), Len(request.servervariables("SCRIPT_NAME")) - 1), "/"))
	 End If


	  If session("URLPORT") = "N" Then
		 XMESAURL = tiposite("") & request.servervariables("SERVER_NAME") & XMESAURL
	  Else
		 XMESAURL = tiposite("") & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL
	  End If
	surx=	XMESAURL
	sHref = surx & "/" & rsx("dirapli") & "/" & Trim(rsx("asp"))
	sTarget = "info"
	sPath = rsx("dirapli")
	If rsx("target") <> "" Then sTarget = rsx("target")
	spt = InStr(rsx("asp"), ".asp")
	xesm = "N"
	If spt > 1 Then
		If Left(rsx("asp"), spt - 1) = "Submenus" Then
			sPath = "../coress20/"
			xesm = "S"
		End If
	End If
	sseg = InStr(sHref, "?")
	If sseg = 0 Then
		sHref = sHref & "?"
	Else
		sHref = sHref & "&"
	End If

	sHref = sHref & "seg=" & OFV.convertircar(sInput, " ", "_") & "&MNN=" & OFV.convertircar(UCase(rsx("menu")), " ", "_")
	if instr(lcase(rsx("asp")),"base=") = 0 then
	  ' agrego base=D para los menus que no son de formularios
	  sHref = sHref &  "&base=D"
	end if

	ssMnn = "MNN" & OFV.convertircar(UCase(rsx("menu")), " ", "_")

	sAccion= sHref
 if nivel=1 then
  if instr(lcase(sAccion),"submenus.asp") > 0 then
   ss7 = ss7  & "<LI class=spc TITLE='" & RSX("titapli")& "'><font color=  #fdfefe   size=1 face=verdana ><a xtarget=info xhref='"& sAccion & "'>&nbsp;&nbsp;&nbsp;<b><u><i>" & UCase(Replace(STEXTO1,"&nbsp;"," ")) & "</a></font></i></u></b>" '"<li>" &  stexto1  & chr(13)
  else
   ss7 = ss7  & "<LI class=spc2 TITLE='" & RSX("titapli")& "'><font color=  #fdfefe   size=1 face=verdana ><a xtitle='" & rsx("titapli") & "'target=info href='"& sAccion & "'><b><i>" & UCase(Replace(STEXTO1,"&nbsp;"," ")) & "</a></font></i></b>" '"<li>" &  stexto1  & chr(13)
  end if

 else
  ss7 = ss7  & "<LI  class=lix  TITLE='" & RSX("titapli")& "' ><font color=  #fdfefe   size=1 face=verdana ><a xtitle='" & rsx("titapli") & "' target=info href='"& sAccion & "'>&nbsp;&nbsp;&nbsp;" & UCase(Replace(STEXTO1,"&nbsp;"," ")) & "</a></font>" '"<li>" &  stexto1  & chr(13)
 end if

 ss7=ss7 & armasubm(rsx("menuatr"),cmm,rsx("menu"),nivel +1,aplic,rsx("id"),UsrGS)
 ss7=ss7 & "</li>"
 end if
 rsx.movenext
 loop
  call ofv.cerrarconsulta(rsx)
 ss7=ss7 & "</UL></div>" & chr(13)
 end if
 armasubm=ss7
end function


Sub ListarM(mdepen,nivel,cmm)
	Dim rs,rs1
	Dim rsoef
	sql="Select Menux.* From Menux, CorAplic where CorAplic.codapli = Menux.mcodapli and CorAplic.CodApli = 'COREM20' AND Menux.mdepen = '" & mdepen & "'"
	'response.write sql & "<br>"

	Set rs = ofv.crearconsultaEx(sql,cmm,1,parametros)

	If  rs.eof Then
		rseof="S"
	Else
		rseof="N"
	End If

	Dim a(100,1) , i
	Dim imenu , iSm
	' Recorre los subdirectorios
	i=0

	Do While Not rs.eof
		'  ReDim Preserve a(i,0)

		sql2="Select top 1 menux.id From Menux, CorAplic where CorAplic.codapli = Menux.mcodapli and CorAplic.CodApli = 'COREM20' AND Menux.mdepen = '" & rs("menu") & "'"
		Set rs2 = ofv.crearconsultaEx(sql2,cmm,1,parametros)

		If Not rs2.eof Then
			If nivel=0 Then
				str=str & "<li><a href='#'><i class='la la-cog'></i><font size=2 color=  #fdfefe   face=verdana>" & UCase(rs("menu")) & "</font></a><h2>" & rs("menu")   & "</h2>" & Chr(13) &"<UL>" & Chr(13)
			Else
				str=str & "<li><a href='#'><i class='la la-cog'></i><font size=2 color= #fdfefe   face=verdana>" & UCase(rs("menu")) & "</font></a><h2>" & rs("menu") & "</h2>" & Chr(13) &"<UL>" & Chr(13)
			End If

			Call ListarM(RS("MENU"),nivel + 1 ,cmm)
			str=str & "</Li>" & Chr(13)
		Else
			xsel=6
			sInput = Trim(rs(xsel))
			sHref = Trim(rs(2))
			sTarget = "info"
			sPath = Session("CorPath")
			If rs(3) <> "" Then sTarget = rs(3)
			spt = InStr(rs(2), ".asp")
			xesm = "N"
			If spt > 1 Then
				If Left(rs(2), spt - 1) = "Submenus" Then
					sPath = "../coress20/"
					xesm = "S"
				End If
			End If
			sseg = InStr(sHref, "?")
			If sseg = 0 Then
				sHref = sHref & "?"
			Else
				sHref = sHref & "&"
			End If
			sHref = sHref & "seg=" & OFV.convertircar(sInput, " ", "_") & "&MNN=" & OFV.convertircar(UCase(rs(1)), " ", "_")
			ssMnn = "MNN" & OFV.convertircar(UCase(rs(1)), " ", "_")
			str=str & "<LI class=cc><font color=  #fdfefe   size=1 face=verdana ><a onclick=" & Chr(34) & "SetMnn('" & ssMnn & "');" & Chr(34)  & " class=xxxx target=info href='" & sHref & "'>" & UCase(RS("MENU")) & "</a></font></LI>"  & Chr(13)
		End If



		rs.movenext

	Loop
	str=str & "</li></UL>" & Chr(13)
	strf=strf &  str
	str=""
End Sub
%>