<div align="center">

## Simple row coloring


</div>

### Description

The fastest way to do alternate colors of rows in your table (from recordset or something)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rammi](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rammi.md)
**Level**          |Intermediate
**User Rating**    |4.6 (32 globes from 7 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__4-1.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rammi-simple-row-coloring__4-6961/archive/master.zip)





### Source Code

<P>&lt;html&gt;<BR>&lt;head&gt;<BR>&lt;title&gt;Alternate rows
color&lt;/title&gt;<BR>&lt;/head&gt;<BR><FONT
color=#000066><STRONG>&lt;style&gt;<BR>.lnTrue {<BR>&nbsp;background :
Gray;<BR>&nbsp;}</STRONG></FONT></P>
<P><FONT color=#000066><STRONG>.lnFalse {<BR>&nbsp;background :
Silver;<BR>&nbsp;}<BR>&lt;/style&gt;<BR></STRONG></FONT>&lt;body&gt;<BR>&lt;table&gt;<BR>&nbsp;&lt;tr&gt;<BR>&nbsp;&lt;th&gt;Test&lt;/th&gt;<BR>&nbsp;&lt;/tr&gt;<BR>&lt;%<BR>Dim
objRs, blnLine</P>
<P>Set objRs = Server.CreateObject("ADODB.RecordSet")<BR>objRs.Open "SOME QUERY
HERE", "Connection string"<BR><BR><BR><FONT color=#000066><STRONG>blnLine =
False</STRONG></FONT></P>
<P>Do Until objRs.EOF<BR>&nbsp;Response.Write <STRONG><FONT
color=#000066>"&lt;tr class=""ln" &amp; blnLine &amp; """&gt;"
</FONT></STRONG>&amp; _<BR>&nbsp;&nbsp;&nbsp;"&lt;td&gt;" &amp; objRs(0) &amp;
"&lt;/td&gt;&lt;/tr&gt;"<BR>&nbsp;<STRONG><FONT color=#000066>blnLine =
Not(blnLine)&nbsp;</FONT></STRONG><BR>&nbsp;objRs.MoveNext<BR>Loop<BR>Set objRs
= Nothing<BR>%&gt;<BR>&lt;/table&gt;<BR>&lt;/body&gt;<BR>&lt;/html&gt;</P>
<P>--------------------------------------------<BR>Blue bold parts are the most
important...</P>

