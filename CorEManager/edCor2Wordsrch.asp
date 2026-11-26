<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->

<%
   response.expires=0



   wfont = 7
   wspacing = 1
   wsize = 1
   wresto = 0

  xitem = request.querystring("xitem")

  set oPrtG = server.createobject("CorWebQueueXPX.WordExec")

  Set oPrtG.Request = Request
  Set oPrtG.Server = Server
  Set oPrtG.Response = Response
  Set oPrtG.Session = Session
  Set oPrtG.Ofv = Ofv

 sMsg = oPrtG.WordSrch(clng(xitem))

         sPath = left(oPrtG.fpath, InStrRev(oPrtG.fpath, "\", len(oPrtG.fpath)-1))
         soutput = replace(oPrtG.foutput,sPath,"")


 if oPrtG.fSts = 2 then
    shtml = "<html><HEAD><title>Cor E-Manager</title>"
    shtml = shtml & "<style>"
    shtml = shtml & "p.tt {font-family:tahoma;font-size:" & wfont & "pt;color:#cccccc;background-color:#114477;}"
    shtml = shtml & "td.to {font-family:tahoma;font-size:" & wfont & "pt;color:#336699;background-color:#f7eedd;border-style:outset;border-width:2;}"
    shtml = shtml & "td.tt {font-family:tahoma;font-size:" & wfont & "pt;color:#ffffff;}"
    shtml = shtml & "td.ot {font-family:tahoma;font-size:" & wfont & "pt;color:#ffd700;}"
    shtml = shtml & "td.gt {font-family:tahoma;font-size:" & wfont & "pt;color:#f7eedd;}"
    shtml = shtml & "td.dt {font-family:tahoma;font-size:" & wfont & "pt;color:#ffd700;}"
    shtml = shtml & "tr.tt {background-color:#000080;}"
    shtml = shtml & "</style><body bgcolor=#ffffff>"
    shtml = shtml & "<hr size=2 width=100% >"
    shtml = shtml & "<br><br><center>"
    shtml = shtml & " <table style={border-style:outset;border-width:2;background-color:#f7eedd;} cellspacing=0 cellpadding=4 width=80% >"
    shtml = shtml & "<tr><td align=center >"
    shtml = shtml & "<font face=tahoma size=2 color=#000080><b>Esquema convertido:</b></font></td></tr>"
    shtml = shtml & "<tr><td align=center >"
    shtml = shtml & "<a href=..\" & soutput & " ><font face=tahoma size=2 color=#336699 ><b><u>" & oPrtG.foutput & "</u></b></font></a>"
    shtml = shtml & "</td></tr>"
    shtml = shtml & "<tr><td align=center >"
    stit = " - cola de conversion nro:" & xitem & " -<br> ingresado: " & oPrtG.fFi & " - finalizado: " & oPrtG.fFf
    shtml = shtml & "<font face=tahoma size=2 color=#993333 >Estado: <b>" & stit & "</b></font>"
    shtml = shtml & "</td></tr></table>"

 elseif oPrtG.fSts = 3 then
    shtml = "<html><HEAD><title>Cor E-Manager</title>"
    shtml = shtml & "<style>"
    shtml = shtml & "p.tt {font-family:tahoma;font-size:" & wfont & "pt;color:#cccccc;background-color:#114477;}"
    shtml = shtml & "td.to {font-family:tahoma;font-size:" & wfont & "pt;color:#336699;background-color:#f7eedd;border-style:outset;border-width:2;}"
    shtml = shtml & "td.tt {font-family:tahoma;font-size:" & wfont & "pt;color:#ffffff;}"
    shtml = shtml & "td.ot {font-family:tahoma;font-size:" & wfont & "pt;color:#ffd700;}"
    shtml = shtml & "td.gt {font-family:tahoma;font-size:" & wfont & "pt;color:#f7eedd;}"
    shtml = shtml & "td.dt {font-family:tahoma;font-size:" & wfont & "pt;color:#ffd700;}"
    shtml = shtml & "tr.tt {background-color:#000080;}"
    shtml = shtml & "</style><body bgcolor=#ffffff>"
    shtml = shtml & "<hr size=2 width=100% >"
    shtml = shtml & "<br><br><center>"
    shtml = shtml & " <table style={border-style:outset;border-width:2;background-color:#f7eedd;} cellspacing=0 cellpadding=4 width=80% >"
    shtml = shtml & "<tr><td align=center >"
    shtml = shtml & "<font face=tahoma size=2 color=#000080><b>Esquema en conversion:</b></font></td></tr>"
    shtml = shtml & "<tr><td align=center >"
    shtml = shtml & "<font face=tahoma size=2 color=#336699 ><b><u>" & oPrtG.foutput & "</u></b></font>"
    shtml = shtml & "</td></tr>"
    shtml = shtml & "<tr><td align=center >"

    stit = " - En cola de conversion nro:" & xitem & " - ingresado: " & oPrtG.fFi & "<br>finalizo con error<br>"


    shtml = shtml & "<font face=tahoma size=2 color=#993333 >Estado: <b>" & stit & "</b></font>"
    shtml = shtml & "</td></tr></table>"

 else
    shtml = "<html><HEAD><title>Cor E-Manager</title><META http-equiv='refresh' content='5' >"
    shtml = shtml & "<style>"
    shtml = shtml & "p.tt {font-family:tahoma;font-size:" & wfont & "pt;color:#cccccc;background-color:#114477;}"
    shtml = shtml & "td.to {font-family:tahoma;font-size:" & wfont & "pt;color:#336699;background-color:#f7eedd;border-style:outset;border-width:2;}"
    shtml = shtml & "td.tt {font-family:tahoma;font-size:" & wfont & "pt;color:#ffffff;}"
    shtml = shtml & "td.ot {font-family:tahoma;font-size:" & wfont & "pt;color:#ffd700;}"
    shtml = shtml & "td.gt {font-family:tahoma;font-size:" & wfont & "pt;color:#f7eedd;}"
    shtml = shtml & "td.dt {font-family:tahoma;font-size:" & wfont & "pt;color:#ffd700;}"
    shtml = shtml & "tr.tt {background-color:#000080;}"
    shtml = shtml & "</style><body bgcolor=#ffffff>"
    shtml = shtml & "<hr size=2 width=100% >"
    shtml = shtml & "<br><br><center>"
    shtml = shtml & " <table style={border-style:outset;border-width:2;background-color:#f7eedd;} cellspacing=0 cellpadding=4 width=80% >"
    shtml = shtml & "<tr><td align=center >"
    shtml = shtml & "<font face=tahoma size=2 color=#000080><b>Esquema en conversion:</b></font></td></tr>"
    shtml = shtml & "<tr><td align=center >"
    shtml = shtml & "<font face=tahoma size=2 color=#336699 ><b><u>" & oPrtG.foutput & "</u></b></font>"
    shtml = shtml & "</td></tr>"
    shtml = shtml & "<tr><td align=center >"
    if oPrtG.fSts = 0 then
       stit = " - En cola de conversion nro:" & xitem & " - ingresado: " & oPrtG.fFi
    else
       stit = " - En proceso de conversion cola nro:" & xitem & " - ingresado: " & oPrtG.fFi
    end if  

    shtml = shtml & "<font face=tahoma size=2 color=#993333 >Estado: <b>" & stit & "</b></font>"
    shtml = shtml & "</td></tr></table>"

 end if

shtml = shtml & "</body></html>"


 
response.write shtml 



 
%>

