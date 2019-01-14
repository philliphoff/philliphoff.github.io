---
layout: posts
title: "Using LINQ with the Office Object Model"
date: 2008-02-19
---
<P>In my last <A href="{% post_url 2008-02-18-query-your-outlook-inbox-with-linq-to-dasl %}">post</A> I talked about LINQ to DASL, a LINQ provider that converts query expressions&nbsp;into their DASL equivalent in order to efficiently filter item collections&nbsp;in Outlook.&nbsp; But LINQ to DASL solves only a very specific problem for one particular application.&nbsp; The Office object model has many types of collections that we might like to use in LINQ expressions.&nbsp; How do we do that?&nbsp; The answer is: it depends.&nbsp; </P>
<P>(If you don't care about the background information, skip to the end of the post to see how&nbsp;you can&nbsp;use the Office Interop API Extensions to simpliy the use of LINQ with the Office object model.)</P>
<P>Most&nbsp;collections in the Office object model implement IEnumerable, which allows them to be used in foreach statements.&nbsp; Let's take Word's Documents collection interface, for example: </P>
<BLOCKQUOTE>
<P><FONT size=2>Word.</FONT><FONT color=#2b91af size=2>Documents</FONT><FONT size=2> docs = Application.Documents;</P></BLOCKQUOTE>
<BLOCKQUOTE>
<P></FONT><FONT color=#0000ff size=2>foreach</FONT><FONT size=2> (Word.</FONT><FONT color=#2b91af size=2>Document</FONT><FONT size=2> doc </FONT><FONT color=#0000ff size=2>in</FONT><FONT size=2> docs)</P>
<P>{</P>
<BLOCKQUOTE>
<P></FONT><FONT color=#2b91af size=2>MessageBox</FONT><FONT size=2>.Show(doc.Name);</P></BLOCKQUOTE>
<P>}</FONT></P></BLOCKQUOTE>
<P>LINQ expressions, however,&nbsp;require collections to implement IEnumerable&lt;T&gt;.&nbsp; Luckily, .NET 3.5 includes the Cast&lt;T&gt;() and OfType() extension methods&nbsp;that&nbsp;convert an IEnumerable into the IEnumerable&lt;T&gt;, as shown in this example:</P>
<BLOCKQUOTE>
<P>&nbsp;<FONT size=2>Word.</FONT><FONT color=#2b91af size=2>Documents</FONT><FONT size=2> docs = Application.Documents;</P>
<P></FONT><FONT color=#0000ff size=2>var</FONT><FONT size=2> names = </FONT></P>
<BLOCKQUOTE>
<P><FONT color=#0000ff size=2>from</FONT><FONT size=2> doc </FONT><FONT color=#0000ff size=2>in</FONT><FONT size=2> docs.Cast&lt;Word.</FONT><FONT color=#2b91af size=2>Document</FONT><FONT size=2>&gt;()</P>
<P></FONT><FONT color=#0000ff size=2>select</FONT><FONT size=2> doc.Name;</P></BLOCKQUOTE>
<P></FONT><FONT color=#0000ff size=2>foreach</FONT><FONT size=2> (</FONT><FONT color=#0000ff size=2>var</FONT><FONT size=2> name </FONT><FONT color=#0000ff size=2>in</FONT><FONT size=2> names)</P>
<P>{</P>
<BLOCKQUOTE>
<P></FONT><FONT color=#2b91af size=2>MessageBox</FONT><FONT size=2>.Show(name);</P></BLOCKQUOTE>
<P>}</P></BLOCKQUOTE></FONT>
<P>There are some collections which do not implement IEnumerable&nbsp;but&nbsp;still expose a GetEnumerator() method.&nbsp; These can&nbsp;also be used in foreach statements.&nbsp; The Windows interface in the Excel object model is one such collection, as shown in this example:</P><FONT size=2>
<BLOCKQUOTE>
<P>Excel.</FONT><FONT color=#2b91af size=2>Windows</FONT><FONT size=2> windows = Application.Windows;</P>
<P></FONT><FONT color=#0000ff size=2>foreach</FONT><FONT size=2> (Excel.</FONT><FONT color=#2b91af size=2>Window</FONT><FONT size=2> window </FONT><FONT color=#0000ff size=2>in</FONT><FONT size=2> windows)</P>
<P>{</P>
<BLOCKQUOTE>
<P></FONT><FONT color=#2b91af size=2>MessageBox</FONT><FONT size=2>.Show(window.WindowNumber.ToString());</P></BLOCKQUOTE>
<P>}</P></BLOCKQUOTE></FONT>
<P>Unfortunately, there is no direct conversion between a type which exposes a&nbsp;GetEnumerator() method and IEnumerable&lt;T&gt;.&nbsp; For that, our own conversion&nbsp;routine is needed.&nbsp; We can actually make this an extension method on the Windows interface to simplify things, as shown below:</P>
<BLOCKQUOTE>
<P>&nbsp;<FONT color=#0000ff size=2>public</FONT><FONT size=2> </FONT><FONT color=#0000ff size=2>static</FONT><FONT size=2> </FONT><FONT color=#0000ff size=2>class</FONT><FONT size=2> </FONT><FONT color=#2b91af size=2>Extensions</P></FONT><FONT size=2>
<P>{</P>
<BLOCKQUOTE>
<P></FONT><FONT color=#0000ff size=2>public</FONT><FONT size=2> </FONT><FONT color=#0000ff size=2>static</FONT><FONT size=2> </FONT><FONT color=#2b91af size=2>IEnumerable</FONT><FONT size=2>&lt;Excel.</FONT><FONT color=#2b91af size=2>Window</FONT><FONT size=2>&gt; ToEnumerable(</FONT><FONT color=#0000ff size=2>this</FONT><FONT size=2> Excel.</FONT><FONT color=#2b91af size=2>Windows</FONT><FONT size=2> windows)</P>
<P>{</P>
<BLOCKQUOTE>
<P></FONT><FONT color=#0000ff size=2>foreach</FONT><FONT size=2> (Excel.</FONT><FONT color=#2b91af size=2>Window</FONT><FONT size=2> window </FONT><FONT color=#0000ff size=2>in</FONT><FONT size=2> windows)</P>
<P>{</P>
<BLOCKQUOTE>
<P></FONT><FONT color=#0000ff size=2>yield</FONT><FONT size=2> </FONT><FONT color=#0000ff size=2>return</FONT><FONT size=2> window;</P></BLOCKQUOTE>
<P>}</P></BLOCKQUOTE>
<P>}</P></BLOCKQUOTE>
<P>}</P></BLOCKQUOTE>
<P>The extension method simply iterates over each element in the collection and yields it back to the caller.&nbsp; (The cast to Window is implicit in the foreach statement.)&nbsp; We can now use the Windows collection in a LINQ expression:</P><FONT size=2>
<BLOCKQUOTE>
<P>Excel.</FONT><FONT color=#2b91af size=2>Windows</FONT><FONT size=2> windows = Application.Windows;</P>
<P></FONT><FONT color=#0000ff size=2>var</FONT><FONT size=2> numbers = </FONT></P>
<BLOCKQUOTE>
<P><FONT color=#0000ff size=2>from</FONT><FONT size=2> window </FONT><FONT color=#0000ff size=2>in</FONT><FONT size=2> windows.ToEnumerable()</P>
<P></FONT><FONT color=#0000ff size=2>select</FONT><FONT size=2> window.WindowNumber.ToString();</P></BLOCKQUOTE>
<P></FONT><FONT color=#0000ff size=2>foreach</FONT><FONT size=2> (</FONT><FONT color=#0000ff size=2>var</FONT><FONT size=2> number </FONT><FONT color=#0000ff size=2>in</FONT><FONT size=2> numbers)</P>
<P>{</P>
<BLOCKQUOTE>
<P></FONT><FONT color=#2b91af size=2>MessageBox</FONT><FONT size=2>.Show(number);</P></BLOCKQUOTE>
<P>}</P></BLOCKQUOTE>
<P>There is one final category of Office collection that does not implement IEnumerable nor expose a GetEnumerator() method.&nbsp; These cannot be used in foreach statements.&nbsp; How then, you might ask, do you iterate over such a collection?&nbsp; The answer is, with a for loop!&nbsp; These collections instead expose a Count (or possibly Length) property and an indexer with which you can&nbsp;enumerate each item.&nbsp; With that we can create an extension method that returns IEnumerable&lt;T&gt; which can be used in a LINQ expression, such as the one shown below:</P><FONT size=2>
<BLOCKQUOTE>
<P></FONT><FONT color=#0000ff size=2>public</FONT><FONT size=2> </FONT><FONT color=#0000ff size=2>static</FONT><FONT size=2> </FONT><FONT color=#2b91af size=2>IEnumerable</FONT><FONT size=2>&lt;</FONT><FONT color=#0000ff size=2>float</FONT><FONT size=2>&gt; ToEnumerable(</FONT><FONT color=#0000ff size=2>this</FONT><FONT size=2> Word.</FONT><FONT color=#2b91af size=2>Adjustments</FONT><FONT size=2> adjustments)</P>
<P>{</P>
<BLOCKQUOTE>
<P></FONT><FONT color=#0000ff size=2>for</FONT><FONT size=2> (</FONT><FONT color=#0000ff size=2>int</FONT><FONT size=2> i = 1; i &lt;= adjustments.Count; i++)</P>
<P>{</P>
<BLOCKQUOTE>
<P></FONT><FONT color=#0000ff size=2>yield</FONT><FONT size=2> </FONT><FONT color=#0000ff size=2>return</FONT><FONT size=2> adjustments[i];</P></BLOCKQUOTE>
<P>}</P></BLOCKQUOTE>
<P>}</P></BLOCKQUOTE>
<P>Note the count from 1 through Count, rather than the typical 0 through Count - 1.&nbsp;&nbsp;The&nbsp;Office object model, originally targeted for languages like VBA, uses 1-based indexing.&nbsp; With that, we can then write the following&nbsp;LINQ expression using the Adjustments interface:</P></FONT><FONT size=2>
<BLOCKQUOTE>
<P>Word.</FONT><FONT color=#2b91af size=2>Shape</FONT><FONT size=2> shape = Shapes.AddTextbox(Microsoft.Office.Core.</FONT><FONT color=#2b91af size=2>MsoTextOrientation</FONT><FONT size=2>.msoTextOrientationHorizontal, 0, 0, 100, 100);</P>
<P>Word.</FONT><FONT color=#2b91af size=2>Adjustments</FONT><FONT size=2> adjustments = shape.Adjustments;</P>
<P></FONT><FONT color=#0000ff size=2>var</FONT><FONT size=2> numbers = </FONT></P>
<BLOCKQUOTE>
<P><FONT color=#0000ff size=2>from</FONT><FONT size=2> adjustment </FONT><FONT color=#0000ff size=2>in</FONT><FONT size=2> adjustments.ToEnumerable()</P>
<P></FONT><FONT color=#0000ff size=2>select</FONT><FONT size=2> adjustment.ToString();</P></BLOCKQUOTE>
<P></FONT><FONT color=#0000ff size=2>foreach</FONT><FONT size=2> (</FONT><FONT color=#0000ff size=2>var</FONT><FONT size=2> number </FONT><FONT color=#0000ff size=2>in</FONT><FONT size=2> numbers)</P>
<P>{</P>
<BLOCKQUOTE>
<P></FONT><FONT color=#2b91af size=2>MessageBox</FONT><FONT size=2>.Show(number);</P></BLOCKQUOTE>
<P>}</P></BLOCKQUOTE>
<P>To sum up, there are three very different ways to use an Office collection in a LINQ expression depending on how each collection interface was defined.&nbsp; If your Office application uses many different collections, that can mean several different coding conventions across your code.&nbsp; Yuck!&nbsp; Not to mention having to remember which collection requires which enumeration method.&nbsp; This is where the Office Interop API Extensions comes in.&nbsp; The Office Interop API Extensions is one of the VSTO Power Tools and simplifies the use of the Office object model.&nbsp; Specifically, it exposes a consistent&nbsp;Items() extension method on&nbsp;many collection interfaces* which return the appropriate IEnumerable&lt;T&gt;.&nbsp; You need not remember to use Cast&lt;T&gt;() for one collection or&nbsp;create your own enumerator for another.&nbsp; Just call Items() and be on your way!&nbsp; Using our previous examples,&nbsp;retrieving an IEnumerable&lt;T&gt; with the Office Interop API Extensions is as simple as:</P><FONT size=2><FONT size=2>
<BLOCKQUOTE>
<P>docs.Items() </FONT><FONT color=#008000 size=2>// Returns IEnumerable&lt;Word.Document&gt;.</FONT></P>
<P>windows.Items() </FONT><FONT color=#008000 size=2>// Returns IEnumerable&lt;Excel.Window&gt;.</FONT></P><FONT size=2>
<P>adjustments.Items() </FONT><FONT color=#008000 size=2>// Returns IEnumerable&lt;float&gt;.</FONT></P></BLOCKQUOTE>
<P>The VSTO Power Tools are expected for release in the very near future.&nbsp;&nbsp;Keep an eye out for them and use the Office Interop API&nbsp;Extensions to enable LINQ in your Office applications!&nbsp;</P>
<P>*Not all of the collections are extended in this initial release.&nbsp; The bulk of the extensions are in the Word and Excel object models, with&nbsp;a key set of collections extended across the rest of the Office suite. </P></FONT></FONT></FONT>