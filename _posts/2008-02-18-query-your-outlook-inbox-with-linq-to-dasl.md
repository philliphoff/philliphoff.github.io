---
layout: post
title: "Query your Outlook Inbox with LINQ to DASL"
date: 2008-02-18
---
<P>Quick, tell me what the following code does:</P>
<BLOCKQUOTE>
<P><FONT size=2>Outlook.</FONT><FONT color=#2b91af size=2>Folder</FONT><FONT size=2> folder = (Outlook.</FONT><FONT color=#2b91af size=2>Folder</FONT><FONT size=2>) </FONT><FONT color=#0000ff size=2>this</FONT><FONT size=2>.Application.Session.GetDefaultFolder(Outlook.</FONT><FONT color=#2b91af size=2>OlDefaultFolders</FONT><FONT size=2>.olFolderInbox);</P>
<P></FONT><FONT color=#0000ff size=2>string</FONT><FONT size=2> subject = </FONT><FONT color=#a31515 size=2>"VSTO"</FONT><FONT size=2>;</P>
<P></FONT><FONT color=#0000ff size=2>string</FONT><FONT size=2> filter = </FONT><FONT color=#a31515 size=2>@"@SQL=(""urn:schemas:httpmail:subject"" LIKE '%"</FONT><FONT size=2> + subject.Replace(</FONT><FONT color=#a31515 size=2>"'"</FONT><FONT size=2>, </FONT><FONT color=#a31515 size=2>"''"</FONT><FONT size=2>) + </FONT><FONT color=#a31515 size=2>@"%' AND ""urn:schemas:httpmail:date"" &lt;= '"</FONT><FONT size=2> + (</FONT><FONT color=#2b91af size=2>DateTime</FONT><FONT size=2>.Now - </FONT><FONT color=#0000ff size=2>new</FONT><FONT size=2> </FONT><FONT color=#2b91af size=2>TimeSpan</FONT><FONT size=2>(7, 0, 0, 0)).ToString(</FONT><FONT color=#a31515 size=2>"g"</FONT><FONT size=2>) + </FONT><FONT color=#a31515 size=2>@"')"</FONT><FONT size=2>;</P>
<P>Outlook.</FONT><FONT color=#2b91af size=2>Items</FONT><FONT size=2> restrictedItems = folder.Items.Restrict(filter);</P>
<P></FONT><FONT color=#0000ff size=2>foreach</FONT><FONT size=2> (Outlook.</FONT><FONT color=#2b91af size=2>MailItem</FONT><FONT size=2> item </FONT><FONT color=#0000ff size=2>in</FONT><FONT size=2> restrictedItems)</P>
<P>{</P>
<BLOCKQUOTE>
<P>System.Diagnostics.</FONT><FONT color=#2b91af size=2>Debug</FONT><FONT size=2>.WriteLine(</FONT><FONT color=#2b91af size=2>String</FONT><FONT size=2>.Format(</FONT><FONT color=#a31515 size=2>"Body: {0}"</FONT><FONT size=2>, item.Body));</P></BLOCKQUOTE>
<P>}</P></BLOCKQUOTE>
<P>If you guessed "queries my Outlook inbox for mail with VSTO in the subject over 1 week old" you'd be correct.&nbsp; Not very pretty, is it?&nbsp; Well, that's the DAV Searching and Locating (DASL) query language for you.&nbsp; DASL is one of the ways to return a filtered view of items in an Outlook folder.&nbsp; The syntax is a variant of SQL, but&nbsp;instead of&nbsp;tables and columns you&nbsp;specify DASL properties which roughly correspond to properties on the various Outlook item interfaces.&nbsp; The previous example generates a query that looks something like:</P>
<P>&nbsp;@SQL=("urn:schemas:httpmail:subject" LIKE '%VSTO%' AND "urn:schemas:httpmail:date" &lt;= '2/11/2008 3:20 PM')&nbsp;</P>
<P>Filtering items using DASL is more efficient than, say, iterating over every item in the collection to identify matching elements.&nbsp; However, DASL is not without its rough edges.&nbsp; For example:</P>
<UL>
<LI>Since queries&nbsp;are&nbsp;simple strings, there&nbsp;is no strong-typing of query elements.&nbsp; Furthermore, there are formatting and escaping&nbsp;rules that must be followed when converting types to their string equivalents.</LI>
<LI>As the complexity of the query grows, so does the complexity of building a valid query string.&nbsp; Imagine creating a query string to search for items with "VSTO" and "Word", but not "VSTO" and "Excel" in the subject, where the item is at least&nbsp;2 weeks old, but not more than 6&nbsp;months.&nbsp; Would you like to maintain such a query?</LI>
<LI>There is no definitive, published&nbsp;mapping between&nbsp;DASL properties and their corresponding Outlook item interface property (if there even is one).&nbsp; Furthermore, while some DASL properties are shared across all Outlook item types, many others are specific to a single type and there is no easy way to identify which are which.&nbsp; Currently, the best way to identify a DASL property is to create a filtered view in Outlook and look at the DASL query that it produces.</LI></UL>
<P>The Office Interop API Extensions, one of the VSTO Power Tools to be released in the very near&nbsp;future, is not just targeted at&nbsp;Office developers using C#.&nbsp; The extensions also include a simple LINQ to DASL implementation that allow you to write queries against the Outlook item collections in the same way you would any other LINQ provider.&nbsp; The LINQ expression is evaulated at runtime,&nbsp;the equivalent DASL query string generated and passed to Outlook.&nbsp; We can now rewrite the query as follows:</P><FONT size=2>
<BLOCKQUOTE>
<P>Outlook.</FONT><FONT color=#2b91af size=2>Folder</FONT><FONT size=2> folder = (Outlook.</FONT><FONT color=#2b91af size=2>Folder</FONT><FONT size=2>) </FONT><FONT color=#0000ff size=2>this</FONT><FONT size=2>.Application.Session.GetDefaultFolder(Outlook.</FONT><FONT color=#2b91af size=2>OlDefaultFolders</FONT><FONT size=2>.olFolderInbox);</P>
<P></FONT><FONT color=#0000ff size=2>string</FONT><FONT size=2> subject = </FONT><FONT color=#a31515 size=2>"VSTO"</FONT><FONT size=2>;</P>
<P></FONT><FONT color=#0000ff size=2>var</FONT><FONT size=2> results = </FONT></P>
<BLOCKQUOTE>
<BLOCKQUOTE>
<P><FONT color=#0000ff size=2>from</FONT><FONT size=2> item </FONT><FONT color=#0000ff size=2>in</FONT><FONT size=2> folder.Items.AsQueryable&lt;</FONT><FONT color=#2b91af size=2>Mail</FONT><FONT size=2>&gt;()</P></BLOCKQUOTE></BLOCKQUOTE>
<BLOCKQUOTE>
<BLOCKQUOTE>
<P></FONT><FONT color=#0000ff size=2>where</FONT><FONT size=2> item.Subject.Contains(subject) &amp;&amp; item.Date &lt;= </FONT><FONT color=#2b91af size=2>DateTime</FONT><FONT size=2>.Now - </FONT><FONT color=#0000ff size=2>new</FONT><FONT size=2> </FONT><FONT color=#2b91af size=2>TimeSpan</FONT><FONT size=2>(7, 0, 0, 0)</P>
<P></FONT><FONT color=#0000ff size=2>select</FONT><FONT size=2> item.Body;</P></BLOCKQUOTE></BLOCKQUOTE>
<P></FONT><FONT color=#0000ff size=2>foreach</FONT><FONT size=2> (</FONT><FONT color=#0000ff size=2>var</FONT><FONT size=2> result </FONT><FONT color=#0000ff size=2>in</FONT><FONT size=2> results)</P>
<P>{</P>
<BLOCKQUOTE>
<P>System.Diagnostics.</FONT><FONT color=#2b91af size=2>Debug</FONT><FONT size=2>.WriteLine(</FONT><FONT color=#2b91af size=2>String</FONT><FONT size=2>.Format(</FONT><FONT color=#a31515 size=2>"Body: {0}"</FONT><FONT size=2>, result));</P></BLOCKQUOTE>
<P>}</P></BLOCKQUOTE></FONT>
<P mce_keep="true">Notice that there are no strings anywhere in the query; no DASL properties to remember and no value formatting to worry about.&nbsp; Everything is strongly-typed.&nbsp; Also notice that we can create a projection on the filtered results.&nbsp; If you only care about the mail bodies, why bother returning the rest of the data?</P>
<P mce_keep="true">You might be wondering about the Mail type in the example.&nbsp; One of the problems with pairing LINQ with Outlook is that the item interfaces are&nbsp;all distinct.&nbsp; They do not inherit from a common&nbsp;base interface&nbsp;despite them sharing&nbsp;many common properties.&nbsp; Therefore, you would not be able to create a query that applies to all items.&nbsp; Furthermore, there are properties on these interfaces which do not have exact DASL equivalents (and vice-versa).&nbsp; This also limits LINQ from using the interfaces directly.&nbsp; Finally, because the interfaces are defined within the Office PIAs and not under our control, there is no effective way to create a static mapping between DASL properties and their equivalent properties on these interfaces.&nbsp; </P>
<P mce_keep="true">The LINQ to DASL implementation instead contains a hierarchy of types that parallel&nbsp;the Outlook item interfaces but which form a proper hierarchy, allowing the creation of generic&nbsp;queries.&nbsp; These types&nbsp;also expose strongly-typed properties which are attributed with their equivalent DASL properties, used to generate the proper query strings.&nbsp; Finally, the LINQ to DASL implementation of IQueryProvider and IQueryable&lt;T&gt; is conveniently exposed to the developer&nbsp;via&nbsp;the AsQueryable&lt;T&gt;()&nbsp;extension method&nbsp;on the Items interface.</P>
<P mce_keep="true">The initial release of the Office Interop API Extensions does not contain mappings for all known DASL properties.&nbsp; In the event that your favorite property is missing, however, there is a mechanism to extend the built-in types&nbsp;to allow strongly-typed queries of any DASL property.&nbsp; There is also a mechanism for extending the built-in types to allow strongly-typed queries of custom user properties&nbsp;of Outlook items.&nbsp; I'll&nbsp;have more details about that&nbsp;in a later post.</P></FONT>