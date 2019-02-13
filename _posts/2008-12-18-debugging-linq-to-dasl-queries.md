---
layout: posts
title: "Debugging LINQ-to-DASL Queries"
date: 2008-12-18
tags: [msdn]
---
<p>When your LINQ-to-DASL queries do not return the results you expect, how do you determine where the problem is?&#160; The issue could be that the query simply doesn't do what you expect.&#160; For example, you could be querying the wrong DASL properties and therefore Outlook returns no (or unexpected) items.&#160; (This is easy to do, as DASL property names can be difficult to identify and map to the equivalent Outlook object model property.)&#160; It could also be that the DASL query syntax itself is incorrect or the LINQ-to-DASL provider did not manage the query results property.</p>  <p>Fortunately, the LINQ-to-DASL provider offers a way to gain a little insight into what it is actually doing.&#160; You can attach a logger to the provider that will be passed the exact DASL string passed to Outlook when the query is executed.</p>  <p>Let's say we want to find all appointments that have been modified in the last week.&#160; That is, where the LastModifiedTime is within the last seven days.&#160; First, we'll define a custom class on which to perform the query.&#160; (We do this because the built-in Appointment query class does not have LastModifiedTime.)</p>  <pre class="code"><span style="color: blue">public class </span><span style="color: #2b91af">MyAppointment </span>: <span style="color: #2b91af">Appointment
</span>{
    [<span style="color: #2b91af">OutlookItemProperty</span>(<span style="color: #a31515">&quot;DAV:getlastmodified&quot;</span>)]
    <span style="color: blue">public </span><span style="color: #2b91af">DateTime </span>LastModifiedTime { <span style="color: blue">get </span>{ <span style="color: blue">return </span>Item.LastModificationTime; } }
}</pre>

<p>Next, we'll create a method that generates the query.</p>

<pre class="code"><span style="color: blue">private </span><span style="color: #2b91af">IEnumerable</span>&lt;Outlook.<span style="color: #2b91af">AppointmentItem</span>&gt; CreateQueryWithoutDebuggingInformation(Outlook.<span style="color: #2b91af">Items </span>items, <span style="color: #2b91af">DateTime </span>oldestDate)
{
    <span style="color: blue">var </span>query = <span style="color: blue">from </span>item <span style="color: blue">in </span>items.AsQueryable&lt;<span style="color: #2b91af">MyAppointment</span>&gt;()
                <span style="color: blue">where </span>item.LastModifiedTime &gt; oldestDate.ToUniversalTime()
                <span style="color: blue">select </span>item.Item;

    <span style="color: blue">return </span>query;
}</pre>

<p>The method takes the items collection on which the query is to be performed and the date we wish to filter by.&#160; Note that there is nothing special about this LINQ-to-DASL query at this point.</p>

<p>Finally, we'll write the method to execute the query and show the results.</p>

<pre class="code"><span style="color: blue">private void </span>ThisAddIn_Startup(<span style="color: blue">object </span>sender, System.<span style="color: #2b91af">EventArgs </span>e)
{
    Outlook.<span style="color: #2b91af">Folder </span>folder = (Outlook.<span style="color: #2b91af">Folder</span>)Application.Session.GetDefaultFolder(Outlook.<span style="color: #2b91af">OlDefaultFolders</span>.olFolderCalendar);
    <span style="color: #2b91af">DateTime </span>oldestDate = <span style="color: #2b91af">DateTime</span>.Now.AddDays(-7);

    <span style="color: blue">var </span>query = CreateQueryWithoutDebuggingInformation(folder.Items, oldestDate);

    <span style="color: blue">foreach </span>(<span style="color: blue">var </span>appt <span style="color: blue">in </span>query)
    {
        <span style="color: blue">if </span>(<span style="color: #2b91af">MessageBox</span>.Show(appt.Subject, <span style="color: #a31515">&quot;Appointment Found!&quot;</span>, <span style="color: #2b91af">MessageBoxButtons</span>.OKCancel) == <span style="color: #2b91af">DialogResult</span>.Cancel)
        {
            <span style="color: blue">break</span>;
        }
    }
}</pre>

<p>The method simply retrieves the Calendar folder, calls the query creation method, and then iterates over the results.&#160; If the query doesn't return any items, or returns items that you do not expect, what do we do next?&#160; The answer is to modify our query by attaching a logger so that we can see the generated DASL.</p>

<p>First, we'll create a simple logger that writes to the debugger output window.</p>

<pre class="code"><span style="color: blue">internal class </span><span style="color: #2b91af">DebuggerWriter </span>: <span style="color: #2b91af">TextWriter
</span>{
    <span style="color: blue">public override </span><span style="color: #2b91af">Encoding </span>Encoding
    {
        <span style="color: blue">get </span>{ <span style="color: blue">throw new </span><span style="color: #2b91af">NotImplementedException</span>(); }
    }

    <span style="color: blue">public override void </span>WriteLine(<span style="color: blue">string </span>value)
    {
        <span style="color: #2b91af">Debug</span>.WriteLine(value);
    }
}</pre>

<p>The logger can be any TextWriter-derived class.&#160; Note that the Encoding property must be overridden, but need not actually be implemented.</p>

<p>Next, we'll create a new query generation method.</p>

<pre class="code"><span style="color: blue">private </span><span style="color: #2b91af">IEnumerable</span>&lt;Outlook.<span style="color: #2b91af">AppointmentItem</span>&gt; CreateQueryWithDebuggingInformation(Outlook.<span style="color: #2b91af">Items </span>items, <span style="color: #2b91af">DateTime </span>oldestDate)
{
    <span style="color: blue">var </span>source = <span style="color: blue">new </span><span style="color: #2b91af">ItemsSource</span>&lt;<span style="color: #2b91af">MyAppointment</span>&gt;(items)
    {
        Log = <span style="color: blue">new </span><span style="color: #2b91af">DebuggerWriter</span>()
    };

    <span style="color: blue">var </span>query = <span style="color: blue">from </span>item <span style="color: blue">in </span>source
                <span style="color: blue">where </span>item.LastModifiedTime &gt; oldestDate.ToUniversalTime()
                <span style="color: blue">select </span>item.Item;

    <span style="color: blue">return </span>query;
}</pre>

<p>Notice that the query itself (the from, where, and select) is virtually identical to the previous query.&#160; The only difference is that we explicitly create a ItemsSource.&#160; (This is what AsQueryable() did implicitly in the previous query.)&#160; We then attach an instance of our logger.</p>

<p>Finally, we update our query execution method.</p>

<pre class="code"><span style="color: blue">private void </span>ThisAddIn_Startup(<span style="color: blue">object </span>sender, System.<span style="color: #2b91af">EventArgs </span>e)
{
    Outlook.<span style="color: #2b91af">Folder </span>folder = (Outlook.<span style="color: #2b91af">Folder</span>)Application.Session.GetDefaultFolder(Outlook.<span style="color: #2b91af">OlDefaultFolders</span>.olFolderCalendar);
    <span style="color: #2b91af">DateTime </span>oldestDate = <span style="color: #2b91af">DateTime</span>.Now.AddDays(-7);

    <span style="color: green">//var query = CreateQueryWithoutDebuggingInformation(folder.Items, oldestDate);
    </span><span style="color: blue">var </span>query = CreateQueryWithDebuggingInformation(folder.Items, oldestDate);

    <span style="color: blue">foreach </span>(<span style="color: blue">var </span>appt <span style="color: blue">in </span>query)
    {
        <span style="color: blue">if </span>(<span style="color: #2b91af">MessageBox</span>.Show(appt.Subject, <span style="color: #a31515">&quot;Appointment Found!&quot;</span>, <span style="color: #2b91af">MessageBoxButtons</span>.OKCancel) == <span style="color: #2b91af">DialogResult</span>.Cancel)
        {
            <span style="color: blue">break</span>;
        }
    }
}</pre>

<p>The only change here was to call the new query generation method.&#160; Now when we run the code in the debugger, we'll see the following in our output window:</p>

<p><a href="/assets/posts/DebugOutput.jpg"><img alt="DebugOutput" src="/assets/posts/DebugOutput_thumb.jpg" width="633" height="267" /></a></p>

<p>You can see the exact DASL filter string sent to Outlook when the query was executed.&#160; You can use this to verify that the DASL property names, syntax, and values all look correct.&#160; You can also compare this filter string to the string built using the advanced filtration dialog in Outlook.</p>

{% include_relative msdn-notice.md %}
