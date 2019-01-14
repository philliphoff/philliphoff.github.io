---
layout: post
title: "LINQ to DASL Walkthrough"
date: 2008-02-25
---
<P>Now that the Office Interop API Extensions have been <A href="{% post_url 2008-02-21-office-interop-api-extensions-now-available %}">released</A>, I thought I would post a complete walkthrough of a simple LINQ to DASL application. Let's start with my fictitious Outlook calendar: </P>
<P mce_keep="true"><IMG style="WIDTH: 725px; HEIGHT: 756px" height=756 src="/assets/posts/AppointmentsView.JPG" width=725></P>
<P>This calendar shows that I have four appointments today. The appointments have been categorized as either "Work" (blue) or "Personal" (green). Suppose I would like to create an Outlook add-in that displays my personal appointments on startup. I will first create a new C#-based Outlook 2007 Add-in project in Visual Studio 2008. </P>
<P mce_keep="true"><IMG style="WIDTH: 697px; HEIGHT: 478px" height=478 src="/assets/posts/NewProject.JPG" width=697></P>
<P>Next I'll add a reference to one of the Outlook extension assemblies, which were installed as part of the Office Interop API Extensions. I'll select version 12.0.0.0 of the assembly because I'm using Outlook 2007. Version 11.0.0.0 of the assembly would be used with Outlook 2003. </P>
<P mce_keep="true"><IMG src="/assets/posts/AddReference.JPG"></P>
<P>Before they can be used, I have to tell the compiler to look for the extensions during build. I'll do that through a set of using statements at the beginning of my source file. </P>
<P mce_keep="true"><IMG src="/assets/posts/AddUsingStatements.JPG"></P>
<P>Two using statements are required: </P>
<UL>
<LI>
<DIV><SPAN style="FONT-SIZE: 10pt; COLOR: blue; FONT-FAMILY: Courier New">using</SPAN><SPAN style="FONT-FAMILY: Consolas"> </SPAN><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New">Microsoft.Office.Interop.Outlook.Extensions;</SPAN><SPAN style="FONT-FAMILY: Consolas"> </SPAN></DIV>
<P>This statement brings in the Items.AsQueryable&lt;T&gt;() extension method that I'll use in my LINQ expression. </P></LI>
<LI>
<DIV><SPAN style="FONT-SIZE: 10pt; COLOR: blue; FONT-FAMILY: Courier New">using</SPAN><SPAN style="FONT-FAMILY: Consolas"> </SPAN><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New">Microsoft.Office.Interop.Outlook.Extensions.Linq;</SPAN><SPAN style="FONT-FAMILY: Consolas"> </SPAN></DIV>
<P>This statement brings in the LINQ to DASL types that form the basis for my LINQ expression. </P></LI></UL>
<P>With that, I can then write my LINQ expression in the startup event handler of the add-in: </P>
<P style="MARGIN-LEFT: 36pt"><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">private</SPAN> <SPAN style="COLOR: blue">void</SPAN> ThisAddIn_Startup(<SPAN style="COLOR: blue">object</SPAN> sender, System.<SPAN style="COLOR: #2b91af">EventArgs</SPAN> e) </SPAN></P>
<P style="MARGIN-LEFT: 36pt"><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New">{ </SPAN></P>
<BLOCKQUOTE>
<P style="MARGIN-LEFT: 36pt"><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New">Outlook.<SPAN style="COLOR: #2b91af">Folder</SPAN> folder = (Outlook.<SPAN style="COLOR: #2b91af">Folder</SPAN>) Application.Session.GetDefaultFolder(Outlook.<SPAN style="COLOR: #2b91af">OlDefaultFolders</SPAN>.olFolderCalendar); </SPAN></P>
<P style="MARGIN-LEFT: 36pt"><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">var</SPAN> appointments = </SPAN></P>
<BLOCKQUOTE>
<P style="MARGIN-LEFT: 36pt"><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">from</SPAN> item <SPAN style="COLOR: blue">in</SPAN> folder.Items.AsQueryable&lt;<SPAN style="COLOR: #2b91af">Appointment</SPAN>&gt;() </SPAN></P>
<P style="MARGIN-LEFT: 36pt"><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">where</SPAN> item.Categories.Contains(<SPAN style="COLOR: #a31515">"Personal"</SPAN>) </SPAN></P>
<P style="MARGIN-LEFT: 36pt"><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">select</SPAN> item.Item; </SPAN></P></BLOCKQUOTE>
<P style="MARGIN-LEFT: 36pt"><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">var</SPAN> builder = <SPAN style="COLOR: blue">new</SPAN> <SPAN style="COLOR: #2b91af">StringBuilder</SPAN>(); </SPAN></P>
<P style="MARGIN-LEFT: 36pt"><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New">builder.AppendLine(<SPAN style="COLOR: #a31515">"Personal Appointments:"</SPAN>); </SPAN></P>
<P style="MARGIN-LEFT: 36pt"><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New">builder.AppendLine(); </SPAN></P>
<P style="MARGIN-LEFT: 36pt"><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">foreach</SPAN> (<SPAN style="COLOR: blue">var</SPAN> appointment <SPAN style="COLOR: blue">in</SPAN> appointments) </SPAN></P>
<P style="MARGIN-LEFT: 36pt"><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New">{ </SPAN></P>
<BLOCKQUOTE>
<P style="MARGIN-LEFT: 36pt"><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New">builder.AppendLine(appointment.Subject); </SPAN></P></BLOCKQUOTE>
<P style="MARGIN-LEFT: 36pt"><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New">} </SPAN></P>
<P style="MARGIN-LEFT: 36pt"><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: #2b91af">MessageBox</SPAN>.Show(builder.ToString()); </SPAN></P></BLOCKQUOTE>
<P style="MARGIN-LEFT: 36pt"><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New">} </SPAN></P>
<P>Let's look at this query more closely. The first clause "<SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">from</SPAN> item <SPAN style="COLOR: blue">in</SPAN> folder.Items.AsQueryable&lt;<SPAN style="COLOR: #2b91af">Appointment</SPAN>&gt;()</SPAN>" uses the new Items.AsQueryable&lt;T&gt; extension method. This extension method simply returns a new instance of the ItemsSource&lt;T&gt; class, which implements the LINQ interfaces IQueryProvider and IQueryable. I know this folder contains only appointments, so I specify the Appointment class for the generic type. The Appointment class is the LINQ to DASL class associated with the AppointmentItem interface in Outlook. If the folder contained a mixture of Outlook item types (such as both appointments and meetings), I would either need to use the more generic OutlookItem class or use the MessageClass property in my query to restrict the types of the items returned. </P>
<P>The second clause "<SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">where</SPAN> item.Categories.Contains(<SPAN style="COLOR: #a31515">"Personal"</SPAN>)</SPAN>" is the "meat" of the query. This is the expression translated into a DASL query string and passed to Outlook. Outlook then returns a collection of Items matching the query string. In this case, I want Outlook to return only items where the categories property contains the string "Personal". The where clause can contain a number of different types of expressions: </P>
<UL>
<LI>The typical set of comparisons (==, !=, &lt;, &lt;=, &gt;=, &gt;, &amp;&amp;, ||) </LI>
<LI>Negation (!) </LI>
<LI>Method calls on properties using String.Contains(), String.StartsWith(), and String.EndsWith() </LI>
<LI>Expressions involving user properties (e.g. item.UserProperties["Foo"].Value) </LI></UL>
<P>The last clause "<SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">select</SPAN> item.Item</SPAN>" specifies what items to return from the query. LINQ to DASL will wrap each item returned by Outlook with an instance of the type specified in the AsQueryable&lt;T&gt;() extension method. That instance can be returned as-is, or a projection on that instance can be returned instead. I want the original AppointmentItem instance returned by Outlook so I specify a simple projection that returns the Item property on the Appointment class. The select clause also determines the ultimate type of the returned data, IEnumerable&lt;Outlook.AppointmentItem&gt; in this case. This is what I iterate over in my foreach loop. </P>
<P>Finally, I can hit 'F5' and see the results. </P>
<P mce_keep="true"><IMG src="/assets/posts/Output.JPG"></P>
<P>Hopefully this helps people get started with LINQ to DASL. (If it doesn't, please let me know what else I can cover to make things more clear.) This sample can be found on Code Gallery <A href="https://code.msdn.microsoft.com/Release/ProjectReleases.aspx?ProjectName=OfficeExtensions&amp;ReleaseId=527" mce_href="https://code.msdn.microsoft.com/Release/ProjectReleases.aspx?ProjectName=OfficeExtensions&amp;ReleaseId=527">here</A>.</P>