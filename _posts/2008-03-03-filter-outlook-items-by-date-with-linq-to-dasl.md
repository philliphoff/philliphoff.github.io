---
layout: posts
title: "Filter Outlook Items by Date with LINQ to DASL"
date: 2008-03-03
---
<P>I received an email over the weekend asking why the following LINQ to DASL query threw an exception: </P>
<BLOCKQUOTE>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New">Outlook.<SPAN style="COLOR: #2b91af">Folder</SPAN> folder = (Outlook.<SPAN style="COLOR: #2b91af">Folder</SPAN>)Application.Session.GetDefaultFolder(Outlook.<SPAN style="COLOR: #2b91af">OlDefaultFolders</SPAN>.olFolderCalendar); </SPAN></P>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">var</SPAN> appointments = </SPAN></P>
<BLOCKQUOTE>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">from</SPAN> item <SPAN style="COLOR: blue">in</SPAN> folder.Items.AsQueryable&lt;<SPAN style="COLOR: #2b91af">Appointment</SPAN>&gt;() </SPAN></P>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">where</SPAN> item.Categories.Contains(<SPAN style="COLOR: #a31515">"Personal Appointments"</SPAN>) &amp;&amp; item.Item.Start.Date &gt;= <SPAN style="COLOR: #2b91af">DateTime</SPAN>.Now - <SPAN style="COLOR: blue">new</SPAN> <SPAN style="COLOR: #2b91af">TimeSpan</SPAN>(30, 0, 0, 0) </SPAN></P>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">select</SPAN> item.Item; </SPAN></P></BLOCKQUOTE>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">foreach</SPAN> (<SPAN style="COLOR: blue">var</SPAN> appointment <SPAN style="COLOR: blue">in</SPAN> appointments) </SPAN></P>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New">{ </SPAN></P>
<BLOCKQUOTE>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: #2b91af">MessageBox</SPAN>.Show(appointment.Start.ToString()); </SPAN></P></BLOCKQUOTE>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New">} </SPAN></P></BLOCKQUOTE>
<P>The query looks simple enough—return all personal appointments for the last 30 days—but when the foreach loop executes a MissingPropertyAttributeException is thrown stating "The property Date on class DateTime does not have an attached OutlookItemUserPropertyAttribute". The problem here is that (as the exception indicates) the Appointment.Item.Start.Date property does not have an OutlookItemProperty or OutlookItemUserProperty attached. These attributes are used by LINQ to DASL to map properties defined on .NET classes to DASL properties defined by Outlook. Why doesn't this attribute exist? Appointment.Item is of type Microsoft.Office.Interop.Outlook.AppointmentItem. This type is part of the Outlook object model and defined by the Outlook PIA. Unfortunately, since we have no control over the Office PIAs, we can't markup the types with our LINQ to DASL attributes. This means that we can't directly query properties on Outlook items. Instead, we query properties on a proxy class (i.e. Appointment) that we do have control over. </P>
<P>But wait a minute…the Appointment class doesn't have a Start property! Yes, unfortunately we weren't able to map every known DASL property to its Outlook item equivalent for this initial release (only so many hours in the day and all that). We did, however, provide a way for you to add such properties yourself. </P>
<BLOCKQUOTE>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">internal</SPAN> <SPAN style="COLOR: blue">class</SPAN> <SPAN style="COLOR: #2b91af">MyAppointment</SPAN> : <SPAN style="COLOR: #2b91af">Appointment </SPAN></SPAN></P>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New">{ </SPAN></P>
<BLOCKQUOTE>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New">[<SPAN style="COLOR: #2b91af">OutlookItemProperty</SPAN>(<SPAN style="COLOR: #a31515">"urn:schemas:calendar:dtstart"</SPAN>)] </SPAN></P>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">public</SPAN> <SPAN style="COLOR: #2b91af">DateTime</SPAN> Start { <SPAN style="COLOR: blue">get</SPAN> { <SPAN style="COLOR: blue">return</SPAN> Item.Start; } } </SPAN></P>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New">[<SPAN style="COLOR: #2b91af">OutlookItemProperty</SPAN>(<SPAN style="COLOR: #a31515">"urn:schemas:calendar:dtend"</SPAN>)] </SPAN></P>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">public</SPAN> <SPAN style="COLOR: #2b91af">DateTime</SPAN> End { <SPAN style="COLOR: blue">get</SPAN> { <SPAN style="COLOR: blue">return</SPAN> Item.End; } } </SPAN></P></BLOCKQUOTE>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New">} </SPAN></P></BLOCKQUOTE>
<P>The MyAppointment class above derives from the existing Appointment class and adds two new properties, Start and End. These properties simply defer to the inner Item's Start and End properties. Each property has an OutlookItemPropertyAttribute attached that maps the property to its corresponding DASL property, which can be found using the handy SQL tab of Outlook's custom filter dialog. Next, the query can be revised as follows: </P>
<BLOCKQUOTE>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New">Outlook.<SPAN style="COLOR: #2b91af">Folder</SPAN> folder = (Outlook.<SPAN style="COLOR: #2b91af">Folder</SPAN>)Application.Session.GetDefaultFolder(Outlook.<SPAN style="COLOR: #2b91af">OlDefaultFolders</SPAN>.olFolderCalendar); </SPAN></P>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">var</SPAN> appointments = </SPAN></P>
<BLOCKQUOTE>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">from</SPAN> item <SPAN style="COLOR: blue">in</SPAN> folder.Items.AsQueryable&lt;<SPAN style="COLOR: #2b91af">MyAppointment</SPAN>&gt;() </SPAN></P>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">where</SPAN> item.Categories.Contains(<SPAN style="COLOR: #a31515">"Personal Appointments"</SPAN>) &amp;&amp; item.Start &gt;= <SPAN style="COLOR: #2b91af">DateTime</SPAN>.Now - <SPAN style="COLOR: blue">new</SPAN> <SPAN style="COLOR: #2b91af">TimeSpan</SPAN>(30, 0, 0, 0) </SPAN></P>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">select</SPAN> item.Item; </SPAN></P></BLOCKQUOTE>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: blue">foreach</SPAN> (<SPAN style="COLOR: blue">var</SPAN> appointment <SPAN style="COLOR: blue">in</SPAN> appointments) </SPAN></P>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New">{ </SPAN></P>
<BLOCKQUOTE>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New"><SPAN style="COLOR: #2b91af">MessageBox</SPAN>.Show(appointment.Start.ToString()); </SPAN></P></BLOCKQUOTE>
<P><SPAN style="FONT-SIZE: 10pt; FONT-FAMILY: Courier New">} </SPAN></P></BLOCKQUOTE>
<P>Note that the AsQueryable&lt;T&gt;() extension method now uses the new MyAppointment type and that its Start property is used instead of Item.Start.Date. Run the query again and Outlook should return a collection of appointments instead of an exception (presuming you have any appointments which match the query).</P>