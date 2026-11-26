<!-- #INCLUDE FILE="coress20/IncFile/CorConect.inc" -->

<%
xx=ARMAM()

response.write xx

Function ARMAM()

    xx=getinfouser(UserLogon())

    wktgrupos = Split(session("SegGrpdsc"), Chr(9))

	For wkti = 0 To UBound(wktgrupos)
		suserinf=suserinf & " " & capitalize(wktgrupos(wkti))
	NEXT



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

	ARMAM= listarM2("0",0,cmm,AppH,UsrGS)



End Function

function  ListarM2(mdepen,nivel,cmm,aplic,UsrgS)
	Dim rs,rs1
	Dim rsoef




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


		 ' response.write sql & "<br>"

	Set rs = ofv.crearconsultaEx(sql,cmm,1,parametros)


  ss=""

	Do While Not rs.eof
		'  ReDim Preserve a(i,0)

        SS=SS & armasubm(rs("menuatr"),cmm,rs(1),1,"'" & rs(2) & "'" ,rs(0),UsrGS,Replace(rs("menu") ," ","&nbsp"))

		rs.movenext

	Loop
	 call ofv.cerrarconsulta(rs)
	ss=ss & "" & Chr(13)
	ListarM2=ss
End function



function armasubm(id,cmm,idx,nivel,aplic,idx2,UsrGS,menuStr)
dim ss7
'dim AUsrG
'dim rsx
'dim XMFI

SQL = "Select max(BB.id) as id, BB.MENU,BB.mcodapli,BB.titapli,max(BB.menuatr) as menuatr,MAX(isnull(  CAST(BB.atrGrupo AS INT)  ,1)) as atrGrupo ,MAX(isnull(CAST(BB.atrUsr AS INT) ,1) )as atrUsr, max(BB.morden) as morden,max(BB.ordenapli )  as ordenapli,max(BB.asp )  as asp"
SQL= SQL & ",max(BB.target) as target,max(BB.dirapli) as dirapli"
sql=sql  & " FROM ( " & chr(13)



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


 Set rsx = ofv.crearconsultaEx(sql,cmm,1,parametros)


 Do While Not rsx.eof



 armasubm=armasubm & "--- " & idx2 & "----" &rsx("id") &  "##"& menuStr&"///" & rsx("menu")&  armasubm(rsx("menuatr"),cmm,rsx("menu"),nivel +1,aplic,rsx("id"),UsrGS,menuStr)

 rsx.movenext
 loop
armasubm=  "<BR>" &armasubm
end function



%>