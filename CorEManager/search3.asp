<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->

<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
 CantRegMostrarx = 11
 CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  
  dim datos(10,50)
  dim posy
  search_phrase=request.querystring("YIDER")
  
  if search_phrase<> "" then
    session("search_phrase")=search_phrase
  else
    search_phrase= session("search_phrase")
  end if
  
  sql= GetResultsFullText(search_phrase)
  set cnbsq = ofv.conectar(strconn8)
  
  public  search_array
  public m_no_results, m_style_br, m_style_title, m_style_you_searched_for, m_style_text, m_style_url, m_style_more_results_text, m_style_more_results_link, m_style_more_results_link_Next, m_trailing_words
    public m_charset, m_rank_exact_match, m_rank_partial_match

  m_trailing_words=10
  
  m_search_style = "font-family:Arial; font-size:10pt; color:#000000"
  m_style_you_searched_for = "font-family:Arial; font-size:10pt; color:#000000"
  m_style_title = "font-size:10pt; font-family:Arial;color:#ff0000;font-weight: bold;text-decoration: underline"
  m_style_text = "font-size:9pt; font-family:Arial"
  m_style_noise_word = "font-size:10pt; font-family:Arial; color:#6f6f6f"
  m_style_url = "font-size:10pt; font-family:Arial; color:green"
  m_style_br = "font-size:10pt; font-family:Arial"
  m_style_more_results_text = "font-size:10pt; font-family:Arial;color:#000000;"
  m_style_more_results_link = "font-size:10pt; font-family:Arial;color:#000000;text-decoration: underline"
  m_style_more_results_link_Next = "font-size:12pt; font-family:Arial; font-weight:bold;color:#000000;text-decoration: underline"


 
  sHTML = ofv.FormHeader(session("FormN"))
  shtml = shtml & CorPnlFnc("N","")
  shtml =shtml & "<script languaje=javascript src='incfile/tooltip.js'></script>"
  Formulario = "" 'ofv.Encabezados("info")
  
  set recordset = ofv.crearconsultaEx(sql,cnbsq,1,parametros)
  
  
  xxx=rstoxml(recordset)
  call leerarr(datos)
  
    if Not recordset.EOF then recordset.move(CantRegMoverx)
    posI = 0
    Do While (Not recordset.EOF) and (posi < CantRegMostrarx)
     	   xincodobj=recordset(3)
           xintipobj=recordset(4)
           xniv1in=recordset(5)
           xniv2in=recordset(6)
            text = recordset(1)
            url = recordset(2)
            
            
            
	              title = Trim(recordset(0))
	    
	              if Len(title) = 0 then
	                title_highlighted = "(Untitled)"
	              else
	    
	                title_highlighted = Highlight(title, search_array)
	    
	                if Len(title_highlighted) = 0 then
	                  title_highlighted = title
	                end if
	    
	              end if


	if count = 1 and Len(search_maintain_search_term) > 0 then
            if InStr(url, "?") = 0 then
              search_maintain_search_term = "?" & search_maintain_search_term
            else
              search_maintain_search_term = "&" & search_maintain_search_term
            end if
          end if
           '   sActcon = sActcon & "&Datos=*smenu=AAAA*lista=3"
           xacceso = "codigo*" & xincodobj & "_tipo*" & xintipobj & "_perfil*OBJARC"
   	   xacceso = replace(xacceso,"_","&")
           xacceso = replace(xacceso,"*","=") & "&NIV1=" & xniv1in & "&NIV2=" & xniv2in
           sActcon="Consultar.asp?DOCMODAL=CL&ACC=XP&" & xacceso & "&Datos=*smenu=PORTEXTO&Ysearch=S"
   
   
             clickn=chr(34) &" AbrirCons('"&sActcon&"','pcd');return false;" & chr(34)
            'url="javascript:void(0)"
             Formulario= Formulario & NewLine(0) & NewLine(0) & "<a  class=xxx onMouseout=""hideddrivetip()"" onMouseover=""ddrivetip('"&armatip(xincodobj)&"')"" onclick="&clickn&" href=""" & url & search_maintain_search_term & """ style=""" & m_style_title & """>" & title_highlighted &"</a>"
          '   response.write NewLine(0) & NewLine(0) & "<br><a  style=""" & m_style_title & """>" & title_highlighted &"</a>"
             extracted = ExtractAndHighlight(text, search_phrase, search_array, m_trailing_words)
              
             if Len(Trim(extracted)) > 0 then
               Formulario= Formulario & NewLine(0) & "<br><span style=""" & m_style_text & """>" & extracted & "</span>"
             end if
             url="<hr>"
             Formulario= Formulario & NewLine(0) & "<br><span style=""" & m_style_url & """>" & url & "</span>"
            Formulario= Formulario & NewLine(0) & "<span style=""" & m_style_br & """></span>"

        recordset.Movenext
        posi=posi+1
      Loop
   '  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","0","0","0","","",Formulario,"si")

   sHTML = sHTML  & Formulario & ofv.botoneraESSX("2",recordset,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)
   sHTML = sHTML  & "</body></html>"
end if


private function Highlight(byval text_to_search, byval word_array_to_find)
    Dim ret

    if MultiByteSearch(YiderSearchTerm()) then
      ret = HighlightMultiByte(text_to_search, word_array_to_find)
    else
      ret = HighlightSingleByte(text_to_search, word_array_to_find)
    end if

    Highlight = ret

  end function


 'highlight all instances of the words in word_array_to_find within text_to_search with a <b> tag
  private function HighlightMultiByte(byval text_to_search, byval word_array_to_find)

    Dim bold_tag, count, find, pos, start

    bold_tag = "<b>"

    for count = 0 to UBound(word_array_to_find)

      pos = -1
      start = 1
      find = word_array_to_find(count)
      pos = InStr(start, text_to_search, find)

      while pos <> 0

        text_to_search = Left(text_to_search, pos - 1) & bold_tag & find & "</b>" & Mid(text_to_search, pos + Len(find))
        start = pos + Len(bold_tag) + 1
        pos = InStr(start, text_to_search, find)

      wend

    next

    HighlightMultiByte = text_to_search

  end function


  'highlight all instances of the words in word_array_to_find within text_to_search with a <b> tag
  private function HighlightSingleByte(byval text_to_search, byval word_array_to_find)
    Dim count, expressionmatch, expressionmatched, highlighted, matches, pos, RegExpObject, str, str_left, str_right, text_to_find

    set RegExpObject = New RegExp
    g_set = g_set + 1

    if not PhraseSearch then
      highlighted = replace(text_to_search, " ", "  ")
      'if adjacent matches are present, they must be separated by double spaces or the regular expression won't find them
    else
      highlighted = text_to_search
    end if

    for count = 0 to UBound(word_array_to_find)

      matches = 0
      'response.Write "<br><br>1 it is " & text_to_search
      'response.Write "<br>it is " & word_array_to_find(count)
      'response.Write "<br>it is " & EscapeRegExpChars(word_array_to_find(count))

      'response.Write "<br>it is " & word_array_to_find(count)
      text_to_find = EscapeRegExpChars(word_array_to_find(count))
      'response.Write "<br>it is " & text_to_find

      'response.Write "<br>highlighted is " & Mid(highlighted, 1, 1)

      With RegExpObject
        .Pattern = WordIsolationPattern(text_to_find)
        .IgnoreCase = true
        .Global = True
      End With

      set expressionmatch = RegExpObject.Execute(highlighted)
      g_set = g_set + 1
      'response.Write "<br>expressionmatch.count is " & expressionmatch.count
      'response.Write "<br>highlighted is " & highlighted

      For Each expressionmatched in expressionmatch

        pos = expressionmatched.FirstIndex + 7*matches

        'response.Write "<br>length is " & expressionmatched.length
        'response.Write "<br>pos is " & pos
        'response.Write "<br>expressionmatched is " & expressionmatched

        str_left = Left(highlighted, pos)
        str = Mid(highlighted, pos + 1, expressionmatched.length)
        str_right = Mid(highlighted, pos + 1 + expressionmatched.length)

        'response.Write "<br>str_left is '" & str_left & "'"
        'response.Write "<br>str is '" & str & "'"
        'response.Write "<br>str_right is '" & str_right & "'"

        if Left(Trim(YiderSearchTerm()), 1) <> "*" and Right(Trim(YiderSearchTerm()), 1) = "*" then
          str = Left(str, 1) & "<b>" & Mid(str, 2, Len(str) - 1) & "</b>"

        elseif Left(Trim(YiderSearchTerm()), 1) = "*" and Right(Trim(YiderSearchTerm()), 1) <> "*" then
          str = "<b>" & Mid(str, 1, Len(str)) & "</b>"

        elseif Left(Trim(YiderSearchTerm()), 1) = "*" and Right(Trim(YiderSearchTerm()), 1) = "*" then
          str = "<b>" & str & "</b>"

        else

          if MultiByteSearch(YiderSearchTerm()) then
            str = "<b>" & str & "</b>"
          else

            if pos = 0 then
            'match at beginning of string
              str = "<b>" & Mid(str, 1, Len(str)) & "</b>"

            else
              str = Left(str, 1) & "<b>" & Mid(str, 2, Len(str) - 1) & "</b>"
            end if

          end if

        end if

         highlighted = str_left & str & str_right

        'response.Write "<br>highlighted is '" & highlighted & "'"

        matches = matches + 1

      Next

      set expressionmatch = Nothing
      g_set = g_set - 1

    next

        'response.Write "<br>highlighted is '" & highlighted & "'"

    set RegExpObject = Nothing
    g_set = g_set - 1

    HighlightSingleByte = highlighted
  end function
    private function GetWords(text_to_search, max_words)
      Dim expressionmatch, RegExpObject, words
  
      set RegExpObject = New RegExp
      g_set = g_set + 1
  
      With RegExpObject
        .Pattern = "\s"
        .IgnoreCase = true
        .Global = True
      End With
  
      set expressionmatch = RegExpObject.Execute(text_to_search)
      g_set = g_set + 1
  
      'response.write "<br>expressionmatch.count is " & expressionmatch.count
      'response.write "<br>max_words is " & max_words
  
      if expressionmatch.count < max_words then
        words = text_to_search
      else
        words = Left(text_to_search, expressionmatch.item(max_words - 1).FirstIndex - 1) & "..."
      end if
  
      set expressionmatch = Nothing
      g_set = g_set - 1
  
      set RegExpObject = Nothing
      g_set = g_set - 1
  
      GetWords = words
  
  end function
  
 private function ExtractAndHighlight(byval text_to_search, byval text_to_find, word_array, max_words)
    Dim extracted

    extracted = Extract(text_to_search, text_to_find, word_array, max_words)

    'response.Write "<br>text_to_search is " & text_to_search
    'response.Write "<br>extracted is '" & extracted & "'"

    if Len(extracted) = 0 and max_words <> -1 then
      extracted = GetWords(text_to_search, 2*max_words)
    else
      extracted = Highlight(extracted, word_array)
    end if

    'response.Write "<br>extracted is '" & Len(Trim(extracted)) & "'"

    ExtractAndHighlight = extracted

  end function
  
 private function NewLine(spaces)
    NewLine = vbCrLf & Space(spaces)
  end function
  
funCtion armatip(idC)
texto=""
 set cn = ofv.conectar(ofv.strconn2)
strsql="select a.temadescri,a.stemadescri,b.*,c.estado AS ESTADOD,b1.codiden from CSRELPATHDESCRIex a "
strsql =strsql & " inner join objetos b1 on(b1.codobjeto= a.cobj)"
strsql =strsql & " inner join objprty b on(b.codobjeto= a.cobj)"
strsql =strsql & " inner join estados c on(c.objid=b.estado) where a.cobj=0" & idC
'RESPONSE.WRITE STRSQL
set recordsetip = CN.Execute(STRSQL)
'response.write recordsetip("autor")
xxx=getinfouser(recordsetip("autor"))

texto="<B><U>TEMA</U> :</b>&nbsp;&nbsp;" & recordsetip("TEMADESCRI") & "<BR><B><U>SUBTEMA</U> :</b>&nbsp;&nbsp;"& recordsetip("STEMADESCRI")& "<BR>"
texto= texto & "<B><U>AUTOR</U>:&nbsp;&nbsp;</B>" & REPLACE(SESSION("SGU_NOMBRE")," ","&nbsp;") '& "&nbsp;<B>SUBTEMA : </b>"& recordsetip("STEMADESCRI")& "<BR>"
texto=  texto & "<br><B><U>ESTADO</U>&nbsp;&nbsp;:</B>" & recordsetip("ESTADOD") & "<br><U><b>VIGENCIA</U> : </b>"& fechadorcor( CDate(CLng("0" & recordsetip("fecvigen")) + clng(cdate("01-01-1900"))),"")& "<BR>"
IF recordsetip("VERSION")="O" THEN
 VERSION="(O)"
ELSE
VERSION="V (" & recordsetip("NROVER") & ")"
END IF
texto=  texto & "<B><U>COD. IDENT.</U> :&nbsp;&nbsp;</B>" & recordsetip("CODIDEN")
texto=  texto & "<BR><B><U>VERSIÓN</U>	:&nbsp;&nbsp;</B>" & VERSION
armatip="<b>" & TEXTO & "</b>"
call ofv.cerrarconsulta(recordsetip)
end function


 function GetResultsFullText( search_phrase)
    Dim count, noise_words, query,  search_string, search_strings, success

    search_strings = Split(search_phrase, " ", -1, 1)

    'query = "select [title], [text], [url],codigo,tipo,solapa,index1 from containstable(Yider, *, '"
    query = "select [title], [text], [url],codigo,tipo,solapa,index1 from containstable(Yider, text, '"
    'response.write query
    '"brown dog" and "lazy"') as ct join Yider as y on ct.[key]=y.[key] order by Rank desc

    count = 0
    for each search_string in search_strings

   
        if count <> 0 then
          query = query & " and """ & search_string & """"
        else
          query = query & """" & search_string & """"
        end if

        count = count + 1

      
       AddToArray search_string, search_array

    next

    query = query & "') as ct join Yider as y on ct.[key]=y.[key] order by Rank desc"
    GetResultsFullText=query
 end function  
 
 
 sub AddToArray(item, byref the_array)
 
   Dim array_size
 
   if not IsArray(the_array) then
     the_array = Array()
   end if
 
   array_size = UBound(the_array) + 1
   ReDim Preserve the_array(array_size)
   the_array(array_size) = item
 
 end sub
 
 
 
 function MultiByteCharSet(charset)
 
   Dim mulitbyte
 
   Select Case charset
   
     Case "ASMO-708" mulitbyte = false
     Case "DOS-720" mulitbyte = false
     Case "iso-8859-6" mulitbyte = false
     Case "x-mac-arabic" mulitbyte = false
     Case "windows-1256" mulitbyte = false
     Case "ibm775" mulitbyte = false
     Case "iso-8859-4" mulitbyte = false
     Case "windows-1257" mulitbyte = false
     Case "ibm852" mulitbyte = false
     Case "iso-8859-2" mulitbyte = false
     Case "x-mac-ce" mulitbyte = false
     Case "windows-1250" mulitbyte = false
     Case "x-Europa" mulitbyte = false
     Case "x-IA5-German" mulitbyte = false
     Case "ibm737" mulitbyte = false
     Case "iso-8859-7" mulitbyte = false
     Case "x-mac-greek" mulitbyte = false
     Case "windows-1253" mulitbyte = false
     Case "ibm869" mulitbyte = false
     Case "DOS-862" mulitbyte = false
     Case "iso-8859-8-i" mulitbyte = false
     Case "iso-8859-8" mulitbyte = false
     Case "x-mac-hebrew" mulitbyte = false
     Case "windows-1255" mulitbyte = false
     Case "x-EBCDIC-Arabic" mulitbyte = false
     Case "x-EBCDIC-CyrillicRussian" mulitbyte = false
     Case "x-EBCDIC-CyrillicSerbianBulgarian" mulitbyte = false
     Case "x-EBCDIC-DenmarkNorway" mulitbyte = false
     Case "x-ebcdic-denmarknorway-euro" mulitbyte = false
     Case "x-EBCDIC-FinlandSweden" mulitbyte = false
     Case "x-ebcdic-finlandsweden-euro" mulitbyte = false
     Case "x-ebcdic-finlandsweden-euro" mulitbyte = false
     Case "x-ebcdic-france-euro" mulitbyte = false
     Case "x-EBCDIC-Germany" mulitbyte = false
     Case "x-ebcdic-germany-euro" mulitbyte = false
     Case "x-EBCDIC-GreekModern" mulitbyte = false
     Case "x-EBCDIC-Greek" mulitbyte = false
     Case "x-EBCDIC-Hebrew" mulitbyte = false
     Case "x-EBCDIC-Icelandic" mulitbyte = false
     Case "x-ebcdic-icelandic-euro" mulitbyte = false
     Case "x-ebcdic-international-euro" mulitbyte = false
     Case "x-EBCDIC-Italy" mulitbyte = false
     Case "x-ebcdic-italy-euro" mulitbyte = false
     Case "X-EBCDIC-Spain" mulitbyte = false
     Case "x-ebcdic-spain-euro" mulitbyte = false
     Case "x-EBCDIC-Thai" mulitbyte = false
     Case "CP1026" mulitbyte = false
     Case "x-EBCDIC-Turkish" mulitbyte = false
     Case "x-EBCDIC-UK" mulitbyte = false
     Case "x-ebcdic-uk-euro" mulitbyte = false
     Case "ebcdic-cp-us" mulitbyte = false
     Case "x-ebcdic-cp-us-euro" mulitbyte = false
     Case "ibm861" mulitbyte = false
     Case "x-mac-icelandic" mulitbyte = false
     Case "iso-8859-3" mulitbyte = false
     Case "iso-8859-15" mulitbyte = false
     Case "x-IA5-Norwegian" mulitbyte = false
     Case "IBM437" mulitbyte = false
     Case "x-IA5-Swedish" mulitbyte = false
     Case "ibm857" mulitbyte = false
     Case "iso-8859-9" mulitbyte = false
     Case "x-mac-turkish" mulitbyte = false
     Case "windows-1254" mulitbyte = false
     Case "us-ascii" mulitbyte = false
     Case "ibm850" mulitbyte = false
     Case "x-IA5" mulitbyte = false
     Case "iso-8859-1" mulitbyte = false
     Case "macintosh" mulitbyte = false
     Case "Windows-1252" mulitbyte = false
 
 
     Case "EUC-CN" mulitbyte = true
     Case "hz-gb-2312" mulitbyte = true
     Case "x-mac-chinesesimp" mulitbyte = true
     Case "big5" mulitbyte = true
     Case "x-Chinese-CNS" mulitbyte = true
     Case "x-Chinese-Eten" mulitbyte = true
     Case "x-mac-chinesetrad" mulitbyte = true
     Case "cp866" mulitbyte = true
     Case "iso-8859-5" mulitbyte = true
     Case "koi8-r" mulitbyte = true
     Case "koi8-u" mulitbyte = true
     Case "x-mac-cyrillic" mulitbyte = true
     Case "windows-1251" mulitbyte = true
     Case "x-EBCDIC-JapaneseAndKana" mulitbyte = true
     Case "x-EBCDIC-JapaneseAndJapaneseLatin" mulitbyte = true
     Case "x-EBCDIC-JapaneseAndUSCanada" mulitbyte = true
     Case "x-EBCDIC-JapaneseKatakana" mulitbyte = true
     Case "x-EBCDIC-KoreanAndKoreanExtended" mulitbyte = true
     Case "x-EBCDIC-KoreanExtended" mulitbyte = true
     Case "CP870" mulitbyte = true
     Case "x-EBCDIC-SimplifiedChinese" mulitbyte = true
     Case "x-EBCDIC-Thai" mulitbyte = true
     Case "x-EBCDIC-TraditionalChinese" mulitbyte = true
     Case "x-iscii-as" mulitbyte = true
     Case "x-iscii-be" mulitbyte = true
     Case "x-iscii-de" mulitbyte = true
     Case "x-iscii-gu" mulitbyte = true
     Case "x-iscii-ka" mulitbyte = true
     Case "x-iscii-ma" mulitbyte = true
     Case "x-iscii-or" mulitbyte = true
     Case "x-iscii-pa" mulitbyte = true
     Case "x-iscii-ta" mulitbyte = true
     Case "x-iscii-te" mulitbyte = true
     Case "euc-jp" mulitbyte = true
     Case "iso-2022-jp" mulitbyte = true
     Case "iso-2022-jp" mulitbyte = true
     Case "csISO2022JP" mulitbyte = true
     Case "x-mac-japanese" mulitbyte = true
     Case "shift_jis" mulitbyte = true
     Case "ks_c_5601-1987" mulitbyte = true
     Case "euc-kr" mulitbyte = true
     Case "iso-2022-kr" mulitbyte = true
     Case "Johab" mulitbyte = true
     Case "x-mac-korean" mulitbyte = true
     Case "windows-874" mulitbyte = true
     Case "unicode" mulitbyte = true
     Case "unicodeFFFE" mulitbyte = true
     Case "utf-7" mulitbyte = true
     Case "utf-8" mulitbyte = true
     Case "gb2312" mulitbyte = true
     Case "windows-1258" mulitbyte = true
 
 	  'return true to be conservative, multi byte charsets have more restrictions than non-multibyte charsets
 	  Case else mulitbyte = true 
 	  
   End Select
   
   'response.Write "MultiByteCharSet is " & charset & " "
   'response.Write mulitbyte
   
   MultiByteCharSet = mulitbyte
     
end function


private function MultiByteSearch(search_phrase)
    Dim count, is_multi, the_asc

    is_multi = false
    
    m_charset=""
    if MultiByteCharSet(m_charset) then

      for count = 1 to Len(search_phrase)
        the_asc = Asc(Mid(search_phrase, count))

        'response.Write "<br>count is " & count & " the_asc is " & the_asc

        if (the_asc < 1 or the_asc > 127) then
          is_multi = true
          Exit For
        end if
      next
    end if

    'response.write "<br>it is " & is_multi

    MultiByteSearch = is_multi

  end function


private function WordIsolationPattern(byval word)

    Dim isolation, original_length

    isolation = m_isolation_vbscript

    'response.Write "<br>1 word is " & word

    word = replace(word, "*", "[^ \t\r\n\v\f\/-]+")

    'response.Write "<br>2 word is " & word

    if Left(word, 1) = "*" and Right(word, 1) <> "*" then
      word = "[^" & m_isolation_vbscript & "]+" & Right(word, Len(word) - 1) & m_isolation_vbscript

    elseif Left(word, 1) <> "*" and Right(word, 1) = "*" then
      word = m_isolation_vbscript & Left(word, Len(word) - 1) & "[^" & m_isolation_vbscript & "]+"

    elseif Left(word, 1) = "*" and Right(word, 1) = "*" then
      word = "[^" & m_isolation_vbscript & "]+" & Mid(word, 2, Len(word) - 2) & "[^" & m_isolation_vbscript & "]+"

    else
      word = m_isolation_vbscript & Left(word, Len(word)) & m_isolation_vbscript

    end if

    'multibyte stuff
    if Len(m_isolation_vbscript) = 0 then
      m_isolation_vbscript = replace(m_isolation_vbscript, "[]", "")
    end if

    WordIsolationPattern = word

  end function

 private function EscapeRegExpChars(byval text_to_find)

    text_to_find = replace(text_to_find, "\", "\\")
    text_to_find = replace(text_to_find, ".", "\.")
    'text_to_find = replace(text_to_find, "*", "\*") ' * is a wildcard and should not be escaped
    text_to_find = replace(text_to_find, "?", "\?")
    text_to_find = replace(text_to_find, "(", "\(")
    text_to_find = replace(text_to_find, ")", "\)")
    text_to_find = replace(text_to_find, "{", "\{")
    text_to_find = replace(text_to_find, "}", "\}")
    text_to_find = replace(text_to_find, "[", "\[")
    text_to_find = replace(text_to_find, "]", "\]")
    text_to_find = replace(text_to_find, "$", "\$")
    text_to_find = replace(text_to_find, "|", "\|")
    text_to_find = replace(text_to_find, "+", "\+")

    'response.Write "<br>text_to_find is " & text_to_find

    EscapeRegExpChars = text_to_find

  end function
  
  'returns 'words_after' words after 'start' bytes from the beginning of the string 'str'
    private function WordsAfter(str, start, words_after)
  	  Dim count, pos, ret
  
  	  'Response.Write("<br> str is '" & str & "' start is " & start & " words_after is " & words_after)
  
      count = 0
  
      if Mid(start, start, 1) <> " " and Right(YiderSearchTerm(), 1) = "*" then
        pos = InStr(start, str, " ")
      else
        pos = start
      end if
  
      if pos <> 0 then
  
  	    pos = InStr(NextNonSpace(str, pos), str, " ")
  
  	    if pos <> 0 then
  		    count = count + 1
  	    end if
  
  	    while(pos <> 0 and count < words_after)
  
          pos = InStr(NextNonSpace(str, pos + 1), str, " ")
  
  		    if pos <> 0 then
  		      count = count + 1
  		    end if
  
  	    wend
  
  	  end if
  
  	  if pos = 0 then
  	    pos = Len(str)
  	  else
  	    ret = " ..."
  	  end if
  
  	  'response.Write "<br><br>start is " & start
  	  'response.Write "<br>pos is " & pos
  	  'response.Write "<br>str is " & str & " Len(str) is " & Len(str)
  
  	  ret = RTrim(Mid(str, start, pos - start + 1)) & ret
  
  	  'response.Write "<br>words after is " & ret
  
  	  WordsAfter = ret
  
    end function
  
  private function PreviousNonSpace(str, pos)
  
      Dim ch, pos_non_space
  
  	  pos_non_space = pos
      ch = " "
  
      'response.write "<br>-pos is " & pos & " ch is " & ch
  
  	  while(ch = " " and pos_non_space > 1)
  
  	    pos_non_space = pos_non_space - 1
  	    ch = Mid(str, pos_non_space, 1)
  
        'response.write "<br>--pos is " & pos & " pos_non_space is " & pos_non_space & " ch is " & ch
  	  wend
  
  	  if pos_non_space < 1 then
  	    pos_non_space = 1
  	  end if
  
      PreviousNonSpace = pos_non_space
  
  end function
    'returns 'words_before' words after 'start' bytes from the beginning of the string 'str'
    private function WordsBefore(str, start, words_before)
      Dim count, pos, prev_non_space, ret
  
      count = 0
  
  	  'Response.Write("<br><br>str is " & str)
  	  'Response.Write("<br>start is " & start)
  	  'Response.Write("<br>words_before is " & words_before)
  
      pos = InStrRev(str, " ", start)
      'InStrRev(str, " ", PreviousNonSpace(str, start))
  
  	  'Response.Write("<br>pos is " & pos)
  
  	  while(pos <> 0 and count < words_before)
  
  	    prev_non_space = PreviousNonSpace(str, pos)
  
  	    'Response.Write("<br>pos is " & pos & " prev_non_space is " & prev_non_space)
  	    pos = InStrRev(str, " ", prev_non_space)
  
  		  if pos <> 0 then
  		    count = count + 1
  		  end if
  
  	  wend
  
  	  if pos = 0 then
  	    pos = 1
  	  end if
  
  	  ret = Mid(str, pos, start - pos)
  
  	  if pos <> 1 then
  	    ret = "... " & ret
  	  end if
  
  	  'Response.Write("<br>pos is " & pos & " start - pos is " & start - pos)
  	  'Response.Write("<br>ret is '" & ret & "'<br><br>")
  	  'Response.Write("<br>start is " & start)
  
  	  WordsBefore  = ret
  
  end function
  
  private function NextNonSpace(str, pos)
  
      Dim ch, pos_non_space
  
  	  pos_non_space = pos
      ch = Mid(str, pos, 1)
  
  	  while(ch = " " and pos_non_space < Len(str))
  
  	    pos_non_space = pos_non_space + 1
  	    ch = Mid(str, pos_non_space, 1)
  
  	  wend
  
      NextNonSpace = pos_non_space
  
  end function

private function ExtractMatchedTextSingleByte(text_to_search, byval text_to_find, byval max_words)

    Dim count, expressionmatch, expressionmatched, extracted, RegExpObject, str, str_left, str_right

    text_to_find = EscapeRegExpChars(text_to_find)

    text_to_find = WordIsolationPattern(text_to_find)

    set RegExpObject = New RegExp
    g_set = g_set + 1

    With RegExpObject
      .Pattern = text_to_find
      'response.Write "<br>pattern is " & text_to_find
      .IgnoreCase = true
      .Global = True
    End With

    set expressionmatch = RegExpObject.Execute(text_to_search)
    g_set = g_set + 1

    'response.Write "<br><br>text_to_search is " & text_to_search
    'response.Write "<br>Pattern is " & RegExpObject.Pattern
    'response.Write "<br>count is " & expressionmatch.Count

    For Each expressionmatched in expressionmatch

      'response.Write "<br>expressionmatched.FirstIndex is " & expressionmatched.FirstIndex

      if max_words <> -1 then
        str_left = WordsBefore(text_to_search, expressionmatched.FirstIndex + 1, max_words)
        str_right = WordsAfter(text_to_search, expressionmatched.FirstIndex + 1 + expressionmatched.length, max_words)
        max_words = -1
      else
        str_left = Left(text_to_search, expressionmatched.FirstIndex)
        str_right = Mid(text_to_search, expressionmatched.FirstIndex + 1 + expressionmatched.length)
      end if

      'response.Write "<br><br>str_left is '" & str_left & "'"
      'response.Write "<br>str_right is '" & str_right & "'"

      str = Mid(text_to_search, expressionmatched.FirstIndex + 1, expressionmatched.length)

      extracted = str_left & str & str_right

      'response.Write "<br><br>extracted is " & extracted

      Exit For

    Next

    'response.Write "<br>extracted is '" & extracted & "'"

    set expressionmatch = Nothing
    g_set = g_set - 1

    set RegExpObject = Nothing
    g_set = g_set - 1

    ExtractMatchedTextSingleByte = extracted
  end function

private function ExtractMatchedTextMultiByte(text_to_search, byval text_to_find, byval max_words)
    Dim count, extracted, pos, str, str_left, str_right, size

    size = 40

    pos = InStr(text_to_search, text_to_find)

    if pos <> 0 then

      if max_words <> -1 then

        if pos <= size then
          str_left = Left(text_to_search, pos - 1)
        else
          str_left = Mid(text_to_search, pos - size, size)

        end if

        str_right = Mid(text_to_search, pos + Len(text_to_find), size)
        max_words = -1

      else
        str_left = Left(text_to_search, pos - 1)
        str_right = Mid(text_to_search, pos + Len(text_to_find))
      end if

      extracted = str_left & Mid(text_to_search, pos, Len(text_to_find)) & str_right

    end if

    'response.Write "<br>extracted is " & extracted

    ExtractMatchedTextMultiByte = extracted

  end function
  
  
  private function ExtractMatchedTextSingleByte(text_to_search, byval text_to_find, byval max_words)
  
      Dim count, expressionmatch, expressionmatched, extracted, RegExpObject, str, str_left, str_right
  
      text_to_find = EscapeRegExpChars(text_to_find)
  
      text_to_find = WordIsolationPattern(text_to_find)
  
      set RegExpObject = New RegExp
      g_set = g_set + 1
  
      With RegExpObject
        .Pattern = text_to_find
        'response.Write "<br>pattern is " & text_to_find
        .IgnoreCase = true
        .Global = True
      End With
  
      set expressionmatch = RegExpObject.Execute(text_to_search)
      g_set = g_set + 1
  
      'response.Write "<br><br>text_to_search is " & text_to_search
      'response.Write "<br>Pattern is " & RegExpObject.Pattern
      'response.Write "<br>count is " & expressionmatch.Count
  
      For Each expressionmatched in expressionmatch
  
        'response.Write "<br>expressionmatched.FirstIndex is " & expressionmatched.FirstIndex
  
        if max_words <> -1 then
          str_left = WordsBefore(text_to_search, expressionmatched.FirstIndex + 1, max_words)
          str_right = WordsAfter(text_to_search, expressionmatched.FirstIndex + 1 + expressionmatched.length, max_words)
          max_words = -1
        else
          str_left = Left(text_to_search, expressionmatched.FirstIndex)
          str_right = Mid(text_to_search, expressionmatched.FirstIndex + 1 + expressionmatched.length)
        end if
  
        'response.Write "<br><br>str_left is '" & str_left & "'"
        'response.Write "<br>str_right is '" & str_right & "'"
  
        str = Mid(text_to_search, expressionmatched.FirstIndex + 1, expressionmatched.length)
  
        extracted = str_left & str & str_right
  
        'response.Write "<br><br>extracted is " & extracted
  
        Exit For
  
      Next
  
      'response.Write "<br>extracted is '" & extracted & "'"
  
      set expressionmatch = Nothing
      g_set = g_set - 1
  
      set RegExpObject = Nothing
      g_set = g_set - 1
  
      ExtractMatchedTextSingleByte = extracted
  end function
  
private function ExtractMatchedText(text_to_search, byval text_to_find, byval max_words)
    Dim ret
   ' response.write MultiByteSearch(YiderSearchTerm())
    if MultiByteSearch(YiderSearchTerm()) then
      ret = ExtractMatchedTextMultiByte(text_to_search, text_to_find, max_words)
    else
      ret = ExtractMatchedTextSingleByte(text_to_search, text_to_find, max_words)
    end if

    ExtractMatchedText = ret

  end function
function YiderSearchTerm()
    Dim yider

    yider = Request("yider")

    if Len(yider) = 0 then
      yider = Request.QueryString("yider")
    end if

    YiderSearchTerm = yider

  session("yider")=yider
  end function
  function SearchComplete(text_found, text_to_search, word_array)
      Dim complete, count, expressionmatch, RegExpObject, text_to_find, total
  
      'response.Write "<br>UBound(word_array) is " & UBound(word_array)
  
      set RegExpObject = New RegExp
      g_set = g_set + 1
      total = 0
  
      for count = 0 to UBound(word_array)
  
        text_to_find = EscapeRegExpChars(word_array(count))
        text_to_find = WordIsolationPattern(text_to_find)
  
        'response.Write "<br><br>text_to_find is " & text_to_find
        'response.Write "<br>text_to_search is " & text_to_search
  
        With RegExpObject
          .Pattern = text_to_find
          .IgnoreCase = true
          .Global = True
        End With
  
        set expressionmatch = RegExpObject.Execute(text_to_search)
  
        'response.Write "<br>expressionmatch.count is " & expressionmatch.count
  
        if expressionmatch.count <> 0 then
          total = total + 1
        else
          Exit For
        end if
  
      next
  
      if UBound(word_array) + 1 = total then
        complete = true
      else
        complete = false
      end if
  
      set RegExpObject = Nothing
      g_set = g_set - 1
  
      SearchComplete = complete
  
  end function
  
private function Extract(byval text_to_search, byval text_to_find, word_array, max_words)
    Dim count, extracted, matched_text

    extracted = ExtractMatchedText(text_to_search, text_to_find, max_words)

   ' response.Write "<br>extracted is '" & text_to_find & "'"
   ' response.Write "<br>text_to_find is '" & text_to_find & "'"
   ' response.Write "<br>text_to_search is '" & text_to_search & "'"
   ' response.Write "<br>word_array is " & word_array(0)
   ' response.Write "<br>word_array is " & word_array(1)

    if Len(extracted) = 0 then
      for count = 0 to UBound(word_array)

        'response.Write "<br><br>count is " & count

        if InStr(lcase(matched_text), lcase(word_array(count))) = 0 then
        'check the word is not already in the string
         
          matched_text = ExtractMatchedText(text_to_search, word_array(count), max_words)
        '  response.write extracted & "222"
          if SearchComplete(matched_text, extracted, word_array) then
            extracted = extracted & matched_text
            Exit For
          elseif Len(extracted) > 0 and Len(matched_text) > 0 then
            extracted = extracted & "<br>"
            extracted = extracted & matched_text
          else
            extracted = extracted & matched_text
          end if

          'response.Write "<br>extracted is " & extracted & "<br><br>"

        end if

      next
    end if

    Extract = extracted
  end function
function rstoxml(rs)
dim mirs

mirs=""'"<?xml version='1.0' ?>"

mirs= mirs & "<docs>"
posy=0
do while (not rs.EOF)

           xincodobj=recordset(3)
           xintipobj=recordset(4)
           xniv1in=recordset(5)
           xniv2in=recordset(6)
            text = recordset(1)
            url = recordset(2)
            
            
            
	              title = Trim(recordset(0))
	    
	              if Len(title) = 0 then
	                title_highlighted = "(Untitled)"
	              else
	    
	                title_highlighted = Highlight(title, search_array)
	    
	                if Len(title_highlighted) = 0 then
	                  title_highlighted = title
	                end if
	    
	              end if


	if count = 1 and Len(search_maintain_search_term) > 0 then
            if InStr(url, "?") = 0 then
              search_maintain_search_term = "?" & search_maintain_search_term
            else
              search_maintain_search_term = "&" & search_maintain_search_term
            end if
          end if
           '   sActcon = sActcon & "&Datos=*smenu=AAAA*lista=3"
           xacceso = "codigo*" & xincodobj & "_tipo*" & xintipobj & "_perfil*OBJARC"
   	   xacceso = replace(xacceso,"_","&")
           xacceso = replace(xacceso,"*","=") & "&NIV1=" & xniv1in & "&NIV2=" & xniv2in
           sActcon="Consultar.asp?DOCMODAL=CL&ACC=XP&" & xacceso & "&Datos=*smenu=PORTEXTO&Ysearch=S"
   
   
             extracted = ExtractAndHighlight(text, search_phrase, search_array, m_trailing_words)
              
             if Len(Trim(extracted)) > 0 then
               extracted= NewLine(0) & "<br><span style=""" & m_style_text & """>" & extracted & "</span>"
             end if 
             
             
datos(0,posy)=  rs("codigo")
datos(1,posy)=  rs("tipo")
datos(2,posy)=  rs("solapa")
datos(3,posy)=  rs("codigo")
datos(4,posy)=  sActcon
datos(5,posy)=  title
datos(6,posy)=  extracted


     
'mirs= mirs &"<doc>" & chr(13)
'mirs= mirs &"<id>" & rs("codigo") & "</id>" '& chr(13)
'mirs= mirs & "<tipo>" & rs("tipo") & "</tipo>"& chr(13)
'mirs= mirs & "<solapa>" & rs("solapa") & "</solapa>"& chr(13)
'mirs= mirs & "<url>" & sActcon & "</url>"& chr(13)
''
'mirs= mirs & "<title>" & title & "</title>"& chr(13)
mirs= mirs & "<extracted>" & extracted & "</extracted>"& chr(13)

'mirs= mirs & "</doc>"'& chr(13)
rs.MoveNext
posy=posy+1
loop

mirs= mirs & "</docs>"
rstoxml=mirs
end function

function leerarr(arr)
'response.write posy
for y=0 to posy
for x=0 to 6
response.write arr(x,y) & "--" 
next
response.write "<br>"
next

end function


Response.Write sHTML

%>