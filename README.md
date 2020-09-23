<div align="center">

## Game Programming in Visual Basic \- Lesson Two


</div>

### Description

This article is lesson two in my mini-series of "Game Programming in Visual Basic". If you like it or even dislike please tell me what was wrong and what was good. PLEASE EVERYONE VOTE, IT ONLY TAKES A FEW SECONDS!!!!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-04-06 20:56:22
**By**             |[Gregory English](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/gregory-english.md)
**Level**          |Beginner
**User Rating**    |5.0 (258 globes from 52 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Games](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/games__1-38.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Game\_Progr69583462002\.zip](https://github.com/Planet-Source-Code/gregory-english-game-programming-in-visual-basic-lesson-two__1-33543/archive/master.zip)





### Source Code

```
<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 9">
<meta name=Originator content="Microsoft Word 9">
<link rel=File-List href="./LessonTwo_files/filelist.xml">
<title>Game Programming in Visual Basic</title>
<!--[if gte mso 9]><xml>
 <o:DocumentProperties>
 <o:Author>MR Ronald W. English</o:Author>
 <o:LastAuthor>MR Ronald W. English</o:LastAuthor>
 <o:Revision>16</o:Revision>
 <o:TotalTime>23</o:TotalTime>
 <o:Created>2002-04-06T20:34:00Z</o:Created>
 <o:LastSaved>2002-04-06T20:56:00Z</o:LastSaved>
 <o:Pages>3</o:Pages>
 <o:Words>802</o:Words>
 <o:Characters>4575</o:Characters>
 <o:Company>English Enterprises Inc,</o:Company>
 <o:Lines>38</o:Lines>
 <o:Paragraphs>9</o:Paragraphs>
 <o:CharactersWithSpaces>5618</o:CharactersWithSpaces>
 <o:Version>9.2720</o:Version>
 </o:DocumentProperties>
</xml><![endif]-->
<style>
<!--
 /* Font Definitions */
@font-face
	{font-family:Wingdings;
	panose-1:5 0 0 0 0 0 0 0 0 0;
	mso-font-charset:2;
	mso-generic-font-family:auto;
	mso-font-pitch:variable;
	mso-font-signature:0 268435456 0 0 -2147483648 0;}
@font-face
	{font-family:Tahoma;
	panose-1:2 11 6 4 3 5 4 4 2 4;
	mso-font-charset:0;
	mso-generic-font-family:swiss;
	mso-font-pitch:variable;
	mso-font-signature:16792199 0 0 0 65791 0;}
@font-face
	{font-family:"Arial Unicode MS";
	panose-1:2 11 6 4 2 2 2 2 2 4;
	mso-font-charset:128;
	mso-generic-font-family:swiss;
	mso-font-pitch:variable;
	mso-font-signature:-1 -369098753 63 0 4129023 0;}
@font-face
	{font-family:"\@Arial Unicode MS";
	mso-font-charset:128;
	mso-generic-font-family:swiss;
	mso-font-pitch:variable;
	mso-font-signature:-1 -369098753 63 0 4129023 0;}
 /* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-parent:"";
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";
	color:windowtext;}
h1
	{mso-style-next:Normal;
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	page-break-after:avoid;
	mso-outline-level:1;
	font-size:11.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Tahoma;
	mso-fareast-font-family:"Times New Roman";
	color:windowtext;
	mso-font-kerning:0pt;
	font-weight:bold;}
h2
	{mso-style-next:Normal;
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	page-break-after:avoid;
	mso-outline-level:2;
	font-size:14.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Tahoma;
	mso-fareast-font-family:"Times New Roman";
	mso-bidi-font-family:"Times New Roman";
	color:red;
	font-weight:bold;}
h3
	{mso-style-next:Normal;
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	page-break-after:avoid;
	mso-outline-level:3;
	font-size:10.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Tahoma;
	mso-fareast-font-family:"Times New Roman";
	mso-bidi-font-family:"Times New Roman";
	color:windowtext;
	font-weight:normal;
	font-style:italic;}
h4
	{mso-style-next:Normal;
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	page-break-after:avoid;
	mso-outline-level:4;
	font-size:10.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Tahoma;
	mso-fareast-font-family:"Times New Roman";
	mso-bidi-font-family:"Times New Roman";
	color:windowtext;
	font-weight:bold;}
p.MsoTitle, li.MsoTitle, div.MsoTitle
	{margin:0in;
	margin-bottom:.0001pt;
	text-align:center;
	mso-pagination:widow-orphan;
	font-size:14.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";
	color:windowtext;
	font-weight:bold;}
p.MsoBodyText, li.MsoBodyText, div.MsoBodyText
	{margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:11.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Tahoma;
	mso-fareast-font-family:"Times New Roman";
	color:windowtext;}
p.MsoSubtitle, li.MsoSubtitle, div.MsoSubtitle
	{margin:0in;
	margin-bottom:.0001pt;
	text-align:center;
	mso-pagination:widow-orphan;
	font-size:11.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";
	color:windowtext;
	font-weight:bold;}
p.MsoBodyText2, li.MsoBodyText2, div.MsoBodyText2
	{margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:10.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:Tahoma;
	mso-fareast-font-family:"Times New Roman";
	mso-bidi-font-family:"Times New Roman";
	color:windowtext;}
a:link, span.MsoHyperlink
	{color:blue;
	text-decoration:underline;
	text-underline:single;}
a:visited, span.MsoHyperlinkFollowed
	{color:purple;
	text-decoration:underline;
	text-underline:single;}
p
	{margin-right:0in;
	mso-margin-top-alt:auto;
	mso-margin-bottom-alt:auto;
	margin-left:0in;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Arial Unicode MS";
	color:black;}
@page Section1
	{size:8.5in 11.0in;
	margin:1.0in 1.25in 1.0in 1.25in;
	mso-header-margin:.5in;
	mso-footer-margin:.5in;
	mso-paper-source:0;}
div.Section1
	{page:Section1;}
 /* List Definitions */
@list l0
	{mso-list-id:236282347;
	mso-list-type:hybrid;
	mso-list-template-ids:1985909076 1508958804 67698713 67698715 67698703 67698713 67698715 67698703 67698713 67698715;}
@list l0:level1
	{mso-level-tab-stop:.75in;
	mso-level-number-position:left;
	margin-left:.75in;
	text-indent:-.25in;}
@list l0:level2
	{mso-level-number-format:alpha-lower;
	mso-level-tab-stop:1.25in;
	mso-level-number-position:left;
	margin-left:1.25in;
	text-indent:-.25in;}
ol
	{margin-bottom:0in;}
ul
	{margin-bottom:0in;}
-->
</style>
<!--[if gte mso 9]><xml>
 <o:shapedefaults v:ext="edit" spidmax="2050"/>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <o:shapelayout v:ext="edit">
 <o:idmap v:ext="edit" data="1"/>
 </o:shapelayout></xml><![endif]-->
</head>
<body lang=EN-US link=blue vlink=purple style='tab-interval:.5in'>
<div class=Section1>
<p class=MsoTitle><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>Game Programming in Visual Basic<o:p></o:p></span></p>
<p class=MsoNormal align=center style='text-align:center'><b><span
style='font-size:11.0pt;mso-bidi-font-size:12.0pt;font-family:Tahoma'>By Greg
English<o:p></o:p></span></b></p>
<p class=MsoNormal><b><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></b></p>
<h2><span style='mso-bidi-font-family:Tahoma'>Introduction<o:p></o:p></span></h2>
<p class=MsoBodyText><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt'>Welcome
to the second of a series of tutorials about “Game Programming in Visual
Basic”. This lesson will get you down into the nitty gritty of the Win32 API.
So go ahead and read on and get coding </span><span style='font-size:10.0pt;
mso-bidi-font-size:12.0pt;font-family:Wingdings;mso-ascii-font-family:Tahoma;
mso-hansi-font-family:Tahoma;mso-char-type:symbol;mso-symbol-font-family:Wingdings'><span
style='mso-char-type:symbol;mso-symbol-font-family:Wingdings'>J</span></span><span
style='font-size:10.0pt;mso-bidi-font-size:12.0pt'>.</span></p>
<p class=MsoNormal><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<h2><span style='mso-bidi-font-family:Tahoma'>Getting Started<o:p></o:p></span></h2>
<p class=MsoBodyText2><span style='mso-bidi-font-family:Tahoma'>In this lesson,
you will learn the techniques of the Win32 API to make a catchy little game for
you and your friends to play. All game programming is are techniques that you
learn and put them together to make the next Quake 3 Engine! We will start off
with good old bitblt. The lesson itself won’t be big, but you can reference my
Sample Project of Asteroids in which I made in 3 hours </span><span
style='font-family:Wingdings;mso-ascii-font-family:Tahoma;mso-hansi-font-family:
Tahoma;mso-bidi-font-family:Tahoma;mso-char-type:symbol;mso-symbol-font-family:
Wingdings'><span style='mso-char-type:symbol;mso-symbol-font-family:Wingdings'>J</span></span><span
style='mso-bidi-font-family:Tahoma'>.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<h2><span style='mso-bidi-font-family:Tahoma'>BitBlt<o:p></o:p></span></h2>
<p class=MsoNormal><b><span style='font-family:Tahoma'>What is BitBlt?<o:p></o:p></span></b></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>BitBlt is the main graphics drawing function for the Win32
GDI, there are others like StretchBlt, but they aren’t really needed here. So
lets take a look at the function<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>Public Declare Function BitBlt Lib &quot;gdi32&quot; (ByVal
<b>hDestDC</b> As Long, ByVal <b>X</b> As Long, ByVal <b>Y</b> As Long, ByVal <b>nWidth</b>
As Long, ByVal <b>nHeight</b> As Long, ByVal <b>hSrcDC</b> As Long, ByVal <b>xSrc</b>
As Long, ByVal <b>ySrc</b> As Long, ByVal <b>dwRop</b> As Long) As Long<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><b><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>hDestDC </span></b><span style='font-size:10.0pt;
mso-bidi-font-size:12.0pt;font-family:Tahoma'>– The destination DC(Device
Context) <i>example: frmMain.hdc/picGame.hdc<o:p></o:p></i></span></p>
<p class=MsoNormal><b><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></b></p>
<p class=MsoNormal><b><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>X,Y </span></b><span style='font-size:10.0pt;mso-bidi-font-size:
12.0pt;font-family:Tahoma'>– The coordinates of where you want the Top/Left
part of the graphics being drawn<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><b><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>nWidth, nHeight </span></b><span style='font-size:10.0pt;
mso-bidi-font-size:12.0pt;font-family:Tahoma'>– The dimensions of the graphic
to be drawn.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><b><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>hSrcDC </span></b><span style='font-size:10.0pt;mso-bidi-font-size:
12.0pt;font-family:Tahoma'>– The source DC from which the graphic comes from.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><b><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>xSrc, ySrc </span></b><span style='font-size:10.0pt;
mso-bidi-font-size:12.0pt;font-family:Tahoma'>– The source coordinates from the
hSrcDC you get the image from(nWidth and nHeight determine xSrc2 and ySrc2)<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><b><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>dwRop </span></b><span style='font-size:10.0pt;mso-bidi-font-size:
12.0pt;font-family:Tahoma'>– The rasterization option. <i>Example: SRCCOPY =
Copy as is, SRCINVERT = Inverts the colors, SRCAND = Copies all but the white,
SRCPAINT = Copies all but the black.</i></span><b><i><span style='font-family:
Tahoma'><o:p></o:p></span></i></b></p>
<h1><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></h1>
<h3><b><span style='mso-bidi-font-family:Tahoma'>Tip For Debugging BitBlt<o:p></o:p></span></b></h3>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>If you haven’t noticed, BitBlt is a function, so it will
return a value. If the value returned is less than or equal to zero, then the
execution of BitBlt has failed. Below is sample code for debugging BitBlt<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>[Code Start]<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>Dim RetVal as long<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>RetVal =
BitBlt(frmMain.hdc,0,0,640,480,picLogo.hdc,0,0,SRCCOPY)<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>If RetVal = 0 Then <o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><span style='mso-tab-count:1'>            </span>MsgBox
“BitBlt has failed!”<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><span style='mso-tab-count:1'>            </span>Exit
Sub/Function<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>End If<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>[Code Stop]<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<h4><span style='font-size:11.0pt;mso-bidi-font-size:12.0pt;mso-bidi-font-family:
Tahoma'>Extra BitBlt Stuff</span><span style='mso-bidi-font-family:Tahoma'><o:p></o:p></span></h4>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<h3><span style='mso-bidi-font-family:Tahoma'>Getting Transparent Blts<o:p></o:p></span></h3>
<p class=MsoBodyText2><span style='mso-bidi-font-family:Tahoma'>Sometimes you
will need to get an image by itself(say a character sprite with a green
background, you would need a Mask for the graphic. A Mask is just a Black and
White picture of the graphic.<o:p></o:p></span></p>
<p class=MsoBodyText2><span style='mso-bidi-font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText2><span style='mso-bidi-font-family:Tahoma'>You would draw
the Mask first using SRCAND, then draw the real graphic EXACTLY AFTER IT, using
SRCINVERT. You can get mask creators off PSC, because I don’t have the time to
make one.<o:p></o:p></span></p>
<p class=MsoBodyText2><span style='mso-bidi-font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>[Code Start]<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>BitBlt frmMain.hdc,0,0,640,480,picLogo.hdc,0,0,SRCAND<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>BitBlt
frmMain.hdc,0,0,640,480,picLogoMask.hdc,0,0,SRCINVERT<o:p></o:p></span></p>
<p class=MsoBodyText2><span style='mso-bidi-font-family:Tahoma'>[Code Stop]<o:p></o:p></span></p>
<p class=MsoBodyText2><span style='mso-bidi-font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<h2><span style='mso-bidi-font-family:Tahoma'>GetAsyncKeyState<o:p></o:p></span></h2>
<h1>What is GetAsyncKeyState?</h1>
<p class=MsoBodyText2>This function allows the programmer to access character
input throughout the program without the use of the default
Form_KeyPress/KeyDown/KeyUp events allowing more versatility I would say.<span
style='mso-bidi-font-family:Tahoma'>Lets take a look at the function, its VERY
VERY VERY simple.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>Public Declare Function GetAsyncKeyState Lib
&quot;user32&quot; (ByVal vKey As Long) As Integer<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><b><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>vKey – </span></b><span style='font-size:10.0pt;mso-bidi-font-size:
12.0pt;font-family:Tahoma'>You insert the key constant here to check its
current state, you can use the basic vbKey constants with this.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>This API is very simple to use.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>[Code Start]<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>Dim btnDown as Boolean<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>btnDown = GetAsyncKeyState(vbKeyDown)<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>If btnDown = True Then ‘//the key is being pressed<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><span style='mso-tab-count:1'>            </span>‘//code
here<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>Else<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><span style='mso-tab-count:1'>            </span>‘//code
here<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>End If<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>[Code Stop]<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>A cool way to use this API can be looked at modEngine.bas
in the Asteroids directory.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<h2><span style='mso-bidi-font-family:Tahoma'>SndPlaySound<o:p></o:p></span></h2>
<h1>What is SndPlaySound?</h1>
<p class=MsoBodyText2>This function is pretty easy to use as well, but at the
same time, it can cause some big problems if the flags given are kinda awkward.
So let’s take a look at this function.</p>
<p class=MsoBodyText2><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText2><span style='mso-bidi-font-family:Tahoma'>Public Declare
Function sndPlaySound Lib &quot;winmm.dll&quot; Alias &quot;sndPlaySoundA&quot;
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long<o:p></o:p></span></p>
<p class=MsoBodyText2><span style='mso-bidi-font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText2><b><span style='mso-bidi-font-family:Tahoma'>LpszSoundName
</span></b><span style='mso-bidi-font-family:Tahoma'>= The filename for the
WAVE sound(must be .WAV sound file)<o:p></o:p></span></p>
<p class=MsoBodyText2><span style='mso-bidi-font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText2><b><span style='mso-bidi-font-family:Tahoma'>UFlags </span></b><span
style='mso-bidi-font-family:Tahoma'>- Flags for the sound when it’s played.<o:p></o:p></span></p>
<p class=MsoBodyText2><span style='mso-bidi-font-family:Tahoma'><span
style="mso-spacerun: yes">    </span><b>SND</b>_<b>ASYNC</b> - &amp;H1 lets you
play a new wav sound, interrupting another<o:p></o:p></span></p>
<p class=MsoBodyText2><span style='mso-bidi-font-family:Tahoma'><span
style="mso-spacerun: yes">    </span><b>SND</b>_<b>LOOP</b> - &amp;H8 loops the
wav sound<o:p></o:p></span></p>
<p class=MsoBodyText2><span style='mso-bidi-font-family:Tahoma'><span
style="mso-spacerun: yes">    </span><b>SND</b>_<b>NODEFAULT</b> - &amp;H2 if
wav file not there, then make sure NOTHING plays<o:p></o:p></span></p>
<p class=MsoBodyText2><span style='mso-bidi-font-family:Tahoma'><span
style="mso-spacerun: yes">    </span><b>SND</b>_<b>SYNC</b> - &amp;H0 no
control to program til wav is done playing<o:p></o:p></span></p>
<p class=MsoBodyText2><span style='mso-bidi-font-family:Tahoma'><span
style="mso-spacerun: yes">    </span><b>SND</b>_<b>NOSTOP</b> - &amp;H10 if a
wav file is already playing then it wont interrupt<o:p></o:p></span></p>
<p class=MsoBodyText2><span style='mso-bidi-font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>[Code Start]<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>sndPlaySound App.Path &amp; “\Audio\Sound.wav”, SND_ASYNC
or SND_NODEFAULT<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>[Code Stop]<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>For some basic subs and functions on using sndPlaySound,
refer to the Asteroids example.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<h2><span style='mso-bidi-font-family:Tahoma'>IntersectRect<o:p></o:p></span></h2>
<h1>What is IntersectRect?</h1>
<p class=MsoBodyText2>This function takes to RECT types and determines whether
they overlap each other. Let’s take a look at this function.</p>
<p class=MsoBodyText2><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText2>Public Declare Function IntersectRect Lib
&quot;user32&quot; (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT)
As Long</p>
<p class=MsoBodyText2><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText2><b>lpDestRect </b>– This RECT will receive the area that
the 2 RECTs crossed over. You would be able to use this RECT for pixel perfect
detection. More on that in a later lesson (maybe…)</p>
<p class=MsoBodyText2><b>lpSrc1Rect</b> – The first source RECT</p>
<p class=MsoBodyText2><b>lpSrc2Rect</b> – The second source RECT</p>
<p class=MsoBodyText2><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>[Code Start]<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>Dim tmpRECT as RECT<br>
Dim PlayerX as Integer, PlayerY As Integer<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>Dim CompX As Integer, CompY as Integer<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>Dim PlayerRect as RECT, CompRect As RECT<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>‘//We are assuming the dimenions of the player are 50x50
and the comp 50x50<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>‘//createrect is a helper function I wrote for creating
rects.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>PlayerRect = CreateRect(PlayerX, PlayerY, PlayerX +50, PlayerY
+ 50)<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>CompRect = CreateRect(CompX,CompY,CompX + 50, CompY + 50)<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>If IntersectRect(tmpRECT,PlayerRect,CompRect) = True Then ‘//there
was an overlap between the 2 rects<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'><span style='mso-tab-count:1'>            </span>‘//code
here<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>End If<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma'>[Code Stop]<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'>Using IntersectRect
can provide a mere decent collision detection like I used in the Asteroids
game. Refer to modEngine.bas for my short Collide function for collision
detection.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
font-family:Tahoma;mso-bidi-font-family:"Times New Roman"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<h2><span style='mso-bidi-font-family:Tahoma'>Other APIs Used<o:p></o:p></span></h2>
<p class=MsoBodyText2>I’m well aware of the other 6 or 7 APIs I used in the
lesson, but if you go to voodoovb.thenexus.bc.ca, there are some good tutorials
on all the DC stuff, they are very good and that’s where I learned from. Or you
can check out a kick-azz VB community at rookscape.com/vbgaming, with lots of
other cool tutorials on such stuff, including some APIs I used.</p>
<p class=MsoBodyText2><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText2><b><span style='font-size:14.0pt;mso-bidi-font-size:12.0pt;
mso-bidi-font-family:Tahoma;color:red'>Conclusion</span></b></p>
<p class=MsoBodyText2>With these simple techniques, you can effectively create
a nice 2d game, better than my Asteroids game I made because I made it in 1 – 2
Hours. You must remember, these are just the techniques NEEDED to create the
game, you gotta learn to put them together by yourself, and when you can
program a cool game(even a simple one), you can probably say to yourself, you
can program anything because games require all the basics of the language like
strings(chars in C++ unless in an array), simple math operations, arrays etc… Until
next time, see ya later <span style='font-family:Wingdings;mso-ascii-font-family:
Tahoma;mso-hansi-font-family:Tahoma;mso-char-type:symbol;mso-symbol-font-family:
Wingdings'><span style='mso-char-type:symbol;mso-symbol-font-family:Wingdings'>J</span></span></p>
<p class=MsoBodyText2><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText2>If you have any questions, comments, or ideas about this
lesson please email me at <a href="mailto:EnglishM1@aol.com">EnglishM1@aol.com</a></p>
<p class=MsoBodyText2><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
</div>
</body>
</html>
```

