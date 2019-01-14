---
layout: post
title: "Banish Missing.Value with the Office Interop API Extensions"
date: 2008-02-15
---
<P>I like VSTO.&nbsp; I like C#.&nbsp; What I don't like is having to write VSTO code in C# like:</P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><FONT face="Times New Roman"><SPAN style="FONT-SIZE: 9pt; COLOR: blue; mso-fareast-language: EN-US; mso-no-proof: yes">object</SPAN><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"> fileName = <SPAN style="COLOR: #a31515">"Test.docx"</SPAN>;<?xml:namespace prefix = o ns = "urn:schemas-microsoft-com:office:office" /><o:p></o:p></SPAN></FONT></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><FONT face="Times New Roman"><SPAN style="FONT-SIZE: 9pt; COLOR: blue; mso-fareast-language: EN-US; mso-no-proof: yes">object</SPAN><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"> missing <SPAN style="mso-spacerun: yes">&nbsp;</SPAN>= System.Reflection.<SPAN style="COLOR: #2b91af">Missing</SPAN>.Value;<o:p></o:p></SPAN></FONT></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><o:p><FONT face="Times New Roman">&nbsp;</FONT></o:p></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman">doc.SaveAs(</FONT></SPAN><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman"><SPAN style="COLOR: blue">ref</SPAN> fileName,</FONT></SPAN><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes"> </SPAN><SPAN style="COLOR: blue">ref</SPAN> missing,&nbsp;</FONT></SPAN><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman"><SPAN style="COLOR: blue">ref</SPAN> missing,&nbsp;</FONT></SPAN><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman"><SPAN style="COLOR: blue">ref</SPAN> missing,</FONT></SPAN><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes"> </SPAN><SPAN style="COLOR: blue">ref</SPAN> missing,</FONT></SPAN><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes"> </SPAN><SPAN style="COLOR: blue">ref</SPAN> missing,</FONT></SPAN><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;</SPAN><SPAN style="COLOR: blue">ref</SPAN> missing,</FONT></SPAN><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;</SPAN><SPAN style="COLOR: blue">ref</SPAN> missing,</FONT></SPAN><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;</SPAN><SPAN style="COLOR: blue">ref</SPAN> missing,</FONT></SPAN><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;</SPAN><SPAN style="COLOR: blue">ref</SPAN> missing,</FONT></SPAN><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;</SPAN><SPAN style="COLOR: blue">ref</SPAN> missing,</FONT></SPAN><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;</SPAN><SPAN style="COLOR: blue">ref</SPAN> missing,</FONT></SPAN><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;</SPAN><SPAN style="COLOR: blue">ref</SPAN> missing,</FONT></SPAN><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;</SPAN><SPAN style="COLOR: blue">ref</SPAN> missing,</FONT></SPAN><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;</SPAN><SPAN style="COLOR: blue">ref</SPAN> missing,</FONT></SPAN><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes"> </SPAN><SPAN style="COLOR: blue">ref</SPAN> missing);</FONT></SPAN></P>
<P>This code creates a copy of a Word document, but it's not very elegant.&nbsp; The SaveAs() method accepts a long list of arguments to tweak the way the file is copied.&nbsp; For&nbsp;a simple copy, however,&nbsp;only the first argument--the new file name--is needed.&nbsp;&nbsp;So why am I stuck passing this "missing" value for each and every omitted argument?&nbsp; The definition of the SaveAs method on the Word Document COM interface defines each argument as&nbsp;optional (see below).&nbsp; This would seem to allow callers&nbsp;to omit the irrelevant arguments.&nbsp;&nbsp;VB developers wouldn't give&nbsp;this another thought.</P>
<P><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US"><FONT face="Times New Roman">HRESULT SaveAs(<o:p></o:p></FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </SPAN>[<SPAN style="COLOR: #0070c0">in, optional</SPAN>] VARIANT* FileName, <o:p></o:p></FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </SPAN>[<SPAN style="COLOR: #0070c0">in, optional</SPAN>] VARIANT* FileFormat, <o:p></o:p></FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </SPAN>[<SPAN style="COLOR: #0070c0">in, optional</SPAN>] VARIANT* LockComments, <o:p></o:p></FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </SPAN>[<SPAN style="COLOR: #0070c0">in, optional</SPAN>] VARIANT* Password, <o:p></o:p></FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </SPAN>[<SPAN style="COLOR: #0070c0">in, optional</SPAN>] VARIANT* AddToRecentFiles, <o:p></o:p></FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </SPAN>[<SPAN style="COLOR: #0070c0">in, optional</SPAN>] VARIANT* WritePassword, <o:p></o:p></FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </SPAN>[<SPAN style="COLOR: #0070c0">in, optional</SPAN>] VARIANT* ReadOnlyRecommended, <o:p></o:p></FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </SPAN>[<SPAN style="COLOR: #0070c0">in, optional</SPAN>] VARIANT* EmbedTrueTypeFonts, <o:p></o:p></FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </SPAN>[<SPAN style="COLOR: #0070c0">in, optional</SPAN>] VARIANT* SaveNativePictureFormat, <o:p></o:p></FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </SPAN>[<SPAN style="COLOR: #0070c0">in, optional</SPAN>] VARIANT* SaveFormsData, <o:p></o:p></FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </SPAN>[<SPAN style="COLOR: #0070c0">in, optional</SPAN>] VARIANT* SaveAsAOCELetter, <o:p></o:p></FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </SPAN>[<SPAN style="COLOR: #0070c0">in, optional</SPAN>] VARIANT* Encoding, <o:p></o:p></FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </SPAN>[<SPAN style="COLOR: #0070c0">in, optional</SPAN>] VARIANT* InsertLineBreaks, <o:p></o:p></FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </SPAN>[<SPAN style="COLOR: #0070c0">in, optional</SPAN>] VARIANT* AllowSubstitutions, <o:p></o:p></FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </SPAN>[<SPAN style="COLOR: #0070c0">in, optional</SPAN>] VARIANT* LineEnding, <o:p></o:p></FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US"><FONT face="Times New Roman"><SPAN style="mso-spacerun: yes">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </SPAN>[<SPAN style="COLOR: #0070c0">in, optional</SPAN>] VARIANT* AddBiDiMarks);</FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US"></SPAN>&nbsp;</P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal">Unfortunately, the C# language and compiler do not comprehend optional arguments.&nbsp; What's worse, unlike the rest of the Office object model, Word interfaces use VARIANT* instead of VARIANT.&nbsp; That is, they are passed by reference rather than by value.&nbsp; This means that, not only does the C# developer have to pass a value for each and every argument, he or she must do so by reference.&nbsp; That means creating an extra object on the stack&nbsp;and then&nbsp;passing it to the method using&nbsp;the ref keyword.&nbsp; How tedious!&nbsp; And it gets even worse;&nbsp;because the values are all passed using objects,&nbsp;we've now lost all of the&nbsp;compile-time advantages of the strongly-typed C# language.&nbsp; How easy would it be to accidentally swap the order of the arguments?</P>
<P>In a perfect world, a simple copy of the document could be created like this:</P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman">doc.SaveAs(<SPAN style="COLOR: #a31515">"Test.docx"</SPAN>);<o:p></o:p></FONT></SPAN></P>
<P>The method would take a strongly-typed string argument by value.&nbsp; And if I wanted to change the format of the newly-saved document, I could write the following:</P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman">doc.SaveAs(<SPAN style="COLOR: #a31515">"Test.html"</SPAN>, WdSaveFormat.wdFormatHTML);</FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman"></FONT></SPAN>&nbsp;</P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes">This method would take the same string argument and another strongly-typed format specifier.&nbsp; </SPAN><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman">We could imagine a series of method overloads&nbsp;that&nbsp;incrementally add&nbsp;to the argument list.&nbsp; But what if I need to specify a disjoint set of arguments?&nbsp; Why&nbsp;can't I simply write:</FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman">&nbsp;<o:p>&nbsp;</P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman">doc.SaveAs(<SPAN style="COLOR: blue">new</SPAN> DocumentSaveAsArgs {<o:p></o:p></FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman"><SPAN style="mso-tab-count: 1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </SPAN>FileName = “Test.docx”,<o:p></o:p></FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman"><SPAN style="mso-tab-count: 1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </SPAN>AddBiDiMarks = false<o:p></o:p></FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"><SPAN style="FONT-SIZE: 9pt; mso-fareast-language: EN-US; mso-no-proof: yes"><FONT face="Times New Roman">});<o:p></o:p></FONT></SPAN></P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal"></o:p></FONT></SPAN>&nbsp;</P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal">The method would take an instance of an arguments class where I could set--using the fancy new C# 3.0 object initializer syntax--only the properties necessary.&nbsp; Furthermore, all the properties would be strongly-typed so that any accidentally swapped arguments could be caught at compile-time and not&nbsp;after the&nbsp;application has been deployed to&nbsp;thousands of clients.&nbsp;</P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal" mce_keep="true">&nbsp;</P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal">Am I just dreaming?&nbsp; Must I be content with typing "ref" over and over again for as long as I develop Office applications?&nbsp; Must I leave my beloved C# for the seductive VB?&nbsp; The answer is a resounding NO!&nbsp; The VSTO Power Tools announced at&nbsp;this week's&nbsp;Office Developer Conference are expected to be released in the very near future.&nbsp; One of those tools is the Office Interop API Extensions, a set of libraries that extend the Office object model and provide a more elegant and consistent API for the C# developer.&nbsp; The three&nbsp;examples above are all possible using the Word extensions shipped as part of this tool.&nbsp; Furthermore, many other interfaces* from across the Office object model have been extended in a similar manner in order&nbsp;to make the lives of C# developers easier.&nbsp; Keep an eye out for this tool and use it to banish Missing.Value (and its close cousin Type.Missing) from your C# Office applications.</P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal" mce_keep="true">&nbsp;</P>
<P class=Code style="MARGIN: 0in 0in 0pt; LINE-HEIGHT: normal">*In this initial release, most&nbsp;of the extension work was&nbsp;focused on Word and&nbsp;Excel.&nbsp; The Outlook extensions had an entirely different focus which I'll discuss in&nbsp;a later&nbsp;post.</P>
<P mce_keep="true">&nbsp;</P>