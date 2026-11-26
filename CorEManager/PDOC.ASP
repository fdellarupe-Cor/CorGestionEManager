<%

 response.addHeader  "Content-Type", "application/msword"
' Response.AddHeader "content-disposition","attachment;filename=projects.doc"





'logo=replace(Session("clilogop"),"../","HTTP://SERVER-20/CORV34DES/")
'oPrtG.RCliNomb=REPLACE(oPrtG.RCliNomb," ","_")
'oPrtG.titdoc=REPLACE(oPrtG.titdoc," ","_")

%>


<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 11">
<meta name=Originator content="Microsoft Word 11">
<link rel=File-List href="./reportesgen_archivos/filelist.xml">
<!--[if gte mso 9]><xml>
 <o:DocumentProperties>
  <o:Author>WinXP</o:Author>
  <o:LastAuthor>WinXP</o:LastAuthor>
  <o:Revision>1</o:Revision>
  <o:TotalTime>0</o:TotalTime>
  <o:Created>2008-05-19T15:41:00Z</o:Created>
  <o:LastSaved>2008-05-19T15:41:00Z</o:LastSaved>
  <o:Pages>1</o:Pages>
  <o:Company>Microsoft</o:Company>
  <o:Lines>1</o:Lines>
  <o:Paragraphs>1</o:Paragraphs>
  <o:Version>11.5606</o:Version>
 </o:DocumentProperties>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <w:WordDocument>
  <w:View>Print</w:View>
  <w:HyphenationZone>21</w:HyphenationZone>
  <w:PunctuationKerning/>
  <w:ValidateAgainstSchemas/>
  <w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>
  <w:IgnoreMixedContent>false</w:IgnoreMixedContent>
  <w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>
  <w:Compatibility>
   <w:BreakWrappedTables/>
   <w:SnapToGridInCell/>
   <w:WrapTextWithPunct/>
   <w:UseAsianBreakRules/>
   <w:DontGrowAutofit/>
  </w:Compatibility>
  <w:BrowserLevel>MicrosoftInternetExplorer4</w:BrowserLevel>
 </w:WordDocument>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <w:LatentStyles DefLockedState="false" LatentStyleCount="156">
 </w:LatentStyles>
</xml><![endif]-->
<style>
<!--
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-parent:"";
	margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
p.MsoHeader, li.MsoHeader, div.MsoHeader
	{margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	tab-stops:center 212.6pt right 425.2pt;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
p.MsoFooter, li.MsoFooter, div.MsoFooter
	{margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	tab-stops:center 212.6pt right 425.2pt;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
 /* Page Definitions */

@page
	{mso-footnote-separator:url("reportesgen_archivos/footer.asp?logo=") fs;
	mso-footnote-continuation-separator:url("reportesgen_archivos/footer.asp?logo=") fcs;
	mso-endnote-separator:url("reportesgen_archivos/footer.asp") es;
	mso-endnote-continuation-separator:url("reportesgen_archivos/footer.asp?logo=") ecs;}
@page Section1

	{size:595.3pt 841.9pt;
	margin:1.0cm 1.0cm 1.0cm 1.0cm;
	mso-header-margin:35.4pt;
	mso-footer-margin:35.4pt;
	mso-even-header:url("reportesgen_archivos/footer.asp?logo=") eh1;
	mso-header:url("reportesgen_archivos/footer.asp?logo=") h1;
	mso-even-footer:url("reportesgen_archivos/footer.asp?logo=") ef1;
	mso-footer:url("reportesgen_archivos/footer.asp?logo=") f1;
	mso-first-header:url("reportesgen_archivos/footer.asp?logo=") fh1;
	mso-first-footer:url("reportesgen_archivos/footer.asp?logo=") ff1;
	mso-paper-source:0;}
div.Section1
	{page:Section1;}

	</style>
	<!--[if gte mso 10]>
	<style>
	 /* Style Definitions */
	 table.MsoNormalTable
		{mso-style-name:"Tabla normal";
		mso-tstyle-rowband-size:0;
		mso-tstyle-colband-size:0;
		mso-style-noshow:yes;
		mso-style-parent:"";
		mso-padding-alt:0cm 5.4pt 0cm 5.4pt;
		mso-para-margin:0cm;
		mso-para-margin-bottom:.0001pt;
		mso-pagination:widow-orphan;
		font-size:10.0pt;
		font-family:"Times New Roman";
		mso-ansi-language:#0400;
		mso-fareast-language:#0400;
		mso-bidi-language:#0400;}
	</style>
	<![endif]-->

	</head>


<body lang=ES style='tab-interval:35.4pt'>

<div class=Section1>


<%
 RESPONSE.WRITE "ServerVariables-----------------------------" & chr(10)

  for each x in Request.ServerVariables
    response.write x & "=" &  Request.ServerVariables(X)  & chr(10)
  next
  RESPONSE.WRITE "SESSION-----------------------------<BR>"& chr(10)
  For Each i in Session.Contents
  Response.Write i & "=" & SESSION(I) & "<BR>" & chr(10)
  Next


%>
</div>

</body>

</html>
