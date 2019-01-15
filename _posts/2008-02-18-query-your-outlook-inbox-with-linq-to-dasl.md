---
layout: posts
title: "Query your Outlook Inbox with LINQ to DASL"
date: 2008-02-18
---
Quick, tell me what the following code does:

```csharp
Outlook.Folder folder = (Outlook.Folder) this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

string subject = "VSTO";

string filter = @"@SQL=(""urn:schemas:httpmail:subject"" LIKE '%" + subject.Replace("'", "''") + @"%' AND ""urn:schemas:httpmail:date"" <= '" + (DateTime.Now - new TimeSpan(7, 0, 0, 0)).ToString("g") + @"')";

Outlook.Items restrictedItems = folder.Items.Restrict(filter);

foreach (Outlook.MailItem item in restrictedItems)
{
    System.Diagnostics.Debug.WriteLine(String.Format("Body: {0}", item.Body));
}
```

If you guessed "queries my Outlook inbox for mail with VSTO in the subject over 1 week old" you'd be correct. Not very pretty, is it? Well, that's the DAV Searching and Locating (DASL) query language for you. DASL is one of the ways to return a filtered view of items in an Outlook folder. The syntax is a variant of SQL, but instead of tables and columns you specify DASL properties which roughly correspond to properties on the various Outlook item interfaces. The previous example generates a query that looks something like:

```sql
@SQL=("urn:schemas:httpmail:subject" LIKE '%VSTO%' AND "urn:schemas:httpmail:date" <= '2/11/2008 3:20 PM')
```

Filtering items using DASL is more efficient than, say, iterating over every item in the collection to identify matching elements. However, DASL is not without its rough edges. For example:

- Since queries are simple strings, there is no strong-typing of query elements. Furthermore, there are formatting and escaping rules that must be followed when converting types to their string equivalents.

 - As the complexity of the query grows, so does the complexity of building a valid query string. Imagine creating a query string to search for items with "VSTO" and "Word", but not "VSTO" and "Excel" in the subject, where the item is at least 2 weeks old, but not more than 6 months. Would you like to maintain such a query?

 - There is no definitive, published mapping between DASL properties and their corresponding Outlook item interface property (if there even is one). Furthermore, while some DASL properties are shared across all Outlook item types, many others are specific to a single type and there is no easy way to identify which are which. Currently, the best way to identify a DASL property is to create a filtered view in Outlook and look at the DASL query that it produces.

The Office Interop API Extensions, one of the VSTO Power Tools to be released in the very near future, is not just targeted at Office developers using C#. The extensions also include a simple LINQ to DASL implementation that allow you to write queries against the Outlook item collections in the same way you would any other LINQ provider. The LINQ expression is evaulated at runtime, the equivalent DASL query string generated and passed to Outlook. We can now rewrite the query as follows:

```csharp
Outlook.Folder folder = (Outlook.Folder) this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

string subject = "VSTO";

var results =
    from item in folder.Items.AsQueryable<Mail>()
    where item.Subject.Contains(subject) && item.Date <= DateTime.Now - new TimeSpan(7, 0, 0, 0)
    select item.Body;

foreach (var result in results)
{
    System.Diagnostics.Debug.WriteLine(String.Format("Body: {0}", result));
}
```

Notice that there are no strings anywhere in the query; no DASL properties to remember and no value formatting to worry about. Everything is strongly-typed. Also notice that we can create a projection on the filtered results. If you only care about the mail bodies, why bother returning the rest of the data?

You might be wondering about the Mail type in the example. One of the problems with pairing LINQ with Outlook is that the item interfaces are all distinct. They do not inherit from a common base interface despite them sharing many common properties. Therefore, you would not be able to create a query that applies to all items. Furthermore, there are properties on these interfaces which do not have exact DASL equivalents (and vice-versa). This also limits LINQ from using the interfaces directly. Finally, because the interfaces are defined within the Office PIAs and not under our control, there is no effective way to create a static mapping between DASL properties and their equivalent properties on these interfaces.

The LINQ to DASL implementation instead contains a hierarchy of types that parallel the Outlook item interfaces but which form a proper hierarchy, allowing the creation of generic queries. These types also expose strongly-typed properties which are attributed with their equivalent DASL properties, used to generate the proper query strings. Finally, the LINQ to DASL implementation of `IQueryProvider` and `IQueryable<T>` is conveniently exposed to the developer via the `AsQueryable<T>()` extension method on the Items interface.

The initial release of the Office Interop API Extensions does not contain mappings for all known DASL properties. In the event that your favorite property is missing, however, there is a mechanism to extend the built-in types to allow strongly-typed queries of any DASL property. There is also a mechanism for extending the built-in types to allow strongly-typed queries of custom user properties of Outlook items. I'll have more details about that in a later post.
