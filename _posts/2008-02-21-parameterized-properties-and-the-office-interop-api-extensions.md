---
layout: post
title: "Parameterized Properties and the Office Interop API Extensions"
date: 2008-02-21
---
<p>One of the disadvantages of C# compared with VB is its lack of support for parameterized properties.  Instead, parameterized properties in C# are exposed as normal method calls prefixed with "get_" and "set_".  This is particularly apparent when working with the Office object model as it exposes many such properties, the majority being indexers on collection interfaces.  To make matters worse, some collection interfaces have array indexers rather than indexer properties, which make indexing inconsistent between collection types.  Since many of the parameters are optional or accept varying types, we lose many of the strong typing benefits of the C# language.  The end result is inconsistent, inelegant, and error prone code.
</p><p>Let's take the Documents collection in Word for example.  That interface defines the Item property with a single variant (i.e. object) argument.  According to the MSDN documentation, that argument value should be an integer or a string.  Retrieving a document by integer index looks something like:
</p><p><span style="font-family:Courier New; font-size:10pt">
			<span style="color:blue">object</span> index = 1;
</span></p><p>
 </p><p><span style="font-family:Courier New; font-size:10pt">            Word.<span style="color:#2b91af">Document</span> doc = Application.Documents.get_Item(<span style="color:blue">ref</span> index);
</span></p><p>
 </p><p><span style="font-family:Courier New; font-size:10pt">
			<span style="color:#2b91af">MessageBox</span>.Show(<span style="color:#a31515">"Name: "</span> + doc.Name);
</span></p><p>
 </p><p>Notice the use of 1 instead of 0; remember that the Office object model uses 1-based indexing.  Also note that we have to pass the index by reference, requiring an extra object on the stack.  This is a particular quirk of the Word object model, but one that makes for even more inelegant code.
</p><p>Retrieving a document by name looks something like:
</p><p><span style="font-family:Courier New; font-size:10pt">
			<span style="color:blue">object</span> index = <span style="color:#a31515">"WordDocument1.docx"</span>;
</span></p><p>
 </p><p><span style="font-family:Courier New; font-size:10pt">            Word.<span style="color:#2b91af">Document</span> doc = Application.Documents.get_Item(<span style="color:blue">ref</span> index);
</span></p><p>
 </p><p><span style="font-family:Courier New; font-size:10pt">
			<span style="color:#2b91af">MessageBox</span>.Show(<span style="color:#a31515">"Name: "</span> + doc.Name);
</span></p><p>
 </p><p>Let's clean up these examples using the Office Interop API Extensions, one of the recently released <a href="http://www.microsoft.com/downloads/details.aspx?FamilyId=46B6BF86-E35D-4870-B214-4D7B72B02BF9&amp;displaylang=en">VSTO Power Tools</a>.  For collections, the libraries expose a set of common Item() extension methods with strongly-typed arguments.  With those we can rewrite our examples:
</p><p><span style="font-family:Courier New; font-size:10pt">            Word.<span style="color:#2b91af">Document</span> doc = Application.Documents.Item(1);
</span></p><p>
 </p><p><span style="font-family:Courier New; font-size:10pt">
			<span style="color:#2b91af">MessageBox</span>.Show(<span style="color:#a31515">"Name: "</span> + doc.Name);
</span></p><p>
 </p><p>And:
</p><p><span style="font-family:Courier New; font-size:10pt">            Word.<span style="color:#2b91af">Document</span> doc = Application.Documents.Item(<span style="color:#a31515">"WordDocument1.docx"</span>);
</span></p><p>
 </p><p><span style="font-family:Courier New; font-size:10pt">
			<span style="color:#2b91af">MessageBox</span>.Show(<span style="color:#a31515">"Name: "</span> + doc.Name);
</span></p><p>
 </p><p>Unfortunately, not all of the collections could be extended in the initial release of the VSTO Power Tools.  You will find the Word and Excel object models to be most complete, with key collections extended across the rest of the Office suite.  (If there are particular collections which you think should be extended in the future, please let us know so that we can prioritize them appropriately.)
</p><p>Indexers are not the only parameterized properties which were extended, however.  Let's take another example from Excel:
</p><p><span style="font-family:Courier New; font-size:10pt">            Excel.<span style="color:#2b91af">Worksheet</span> sheet = InnerObject;
</span></p><p><span style="font-family:Courier New; font-size:10pt">            Excel.<span style="color:#2b91af">Range</span> range1 = sheet.get_Range(<span style="color:#a31515">"A1:B1"</span>, System.<span style="color:#2b91af">Type</span>.Missing);
</span></p><p><span style="font-family:Courier New; font-size:10pt">            Excel.<span style="color:#2b91af">Range</span> range2 = sheet.get_Range(<span style="color:#a31515">"A1"</span>, <span style="color:#a31515">"B2"</span>);
</span></p><p>
 </p><p>In this example, we're retrieving two ranges from an Excel Worksheet.  The Range property (as defined by the COM interface) accepts two parameters, both of which are of variant type and the second of which is also marked optional.  Since C# doesn't support optional parameters, we have to pass System.Type.Missing to indicate the absence of the second value.  Since both values are variant types, we do not have any compile-time type checking.  Again, we can use the extensions provided by the Office Interop API Extensions to improve the code:
</p><p><span style="font-family:Courier New; font-size:10pt">            Excel.<span style="color:#2b91af">Worksheet</span> sheet = InnerObject;
</span></p><p><span style="font-family:Courier New; font-size:10pt">            Excel.<span style="color:#2b91af">Range</span> range1 = sheet.Range(<span style="color:#a31515">"A1:B1"</span>);
</span></p><p><span style="font-family:Courier New; font-size:10pt">            Excel.<span style="color:#2b91af">Range</span> range2 = sheet.Range(<span style="color:#a31515">"A1"</span>, <span style="color:#a31515">"B2"</span>);
</span></p><p>
 </p><p>Not only have we eliminated the ugly "get_" and the need to specify System.Type.Missing, our arguments are conveniently strongly-typed.  These may seem like little things, but over the long term I think the improved readability and additional compile-time support will make for a better Office development experience for the C# developer.
</p>