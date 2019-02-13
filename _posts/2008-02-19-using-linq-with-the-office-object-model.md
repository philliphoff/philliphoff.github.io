---
layout: posts
title: "Using LINQ with the Office Object Model"
date: 2008-02-19
tags: [msdn]
---
In my last [post]({% post_url 2008-02-18-query-your-outlook-inbox-with-linq-to-dasl %}) I talked about LINQ to DASL, a LINQ provider that converts query expressions into their DASL equivalent in order to efficiently filter item collections in Outlook. But LINQ to DASL solves only a very specific problem for one particular application. The Office object model has many types of collections that we might like to use in LINQ expressions. How do we do that? The answer is: it depends.

(If you don't care about the background information, skip to the end of the post to see how you can use the Office Interop API Extensions to simpliy the use of LINQ with the Office object model.)

Most collections in the Office object model implement `IEnumerable`, which allows them to be used in `foreach` statements. Let's take Word's `Documents` collection interface, for example:

```csharp
Word.Documents docs = Application.Documents;

foreach (Word.Document doc in docs)
{
    MessageBox.Show(doc.Name);
}
```

LINQ expressions, however, require collections to implement `IEnumerable<T>.` Luckily, .NET 3.5 includes the `Cast<T>()` and `OfType()` extension methods that convert an `IEnumerable` into the `IEnumerable<T>`, as shown in this example:

```csharp
Word.Documents docs = Application.Documents;

var names =
    from doc in docs.Cast<Word.Document>()
    select doc.Name;

foreach (var name in names)
{
    MessageBox.Show(name);
}
```

There are some collections which do not implement `IEnumerable` but still expose a `GetEnumerator()` method. These can also be used in foreach statements. The `Windows` interface in the Excel object model is one such collection, as shown in this example:

```csharp
Excel.Windows windows = Application.Windows;

foreach (Excel.Window window in windows)
{
    MessageBox.Show(window.WindowNumber.ToString());
}
```

Unfortunately, there is no direct conversion between a type which exposes a `GetEnumerator()` method and `IEnumerable<T>`. For that, our own conversion routine is needed. We can actually make this an extension method on the `Windows` interface to simplify things, as shown below:

```csharp
public static class Extensions
{
    public static IEnumerable<Excel.Window> ToEnumerable(this Excel.Windows windows)
    {
        foreach (Excel.Window window in windows)
        {
            yield return window;
        }
    }
}
```

The extension method simply iterates over each element in the collection and yields it back to the caller. (The cast to `Window` is implicit in the `foreach` statement.) We can now use the `Windows` collection in a LINQ expression:

```csharp
Excel.Windows windows = Application.Windows;

var numbers =
    from window in windows.ToEnumerable()
    select window.WindowNumber.ToString();

foreach (var number in numbers)
{
    MessageBox.Show(number);
}
```

There is one final category of Office collection that does not implement `IEnumerable` nor expose a `GetEnumerator()` method. These cannot be used in `foreach` statements. How then, you might ask, do you iterate over such a collection? The answer is, with a `for` loop! These collections instead expose a `Count` (or possibly `Length`) property and an indexer with which you can enumerate each item. With that we can create an extension method that returns `IEnumerable<T>` which can be used in a LINQ expression, such as the one shown below:

Note the count from 1 through `Count`, rather than the typical 0 through `Count - 1`. The Office object model, originally targeted for languages like VBA, uses 1-based indexing. With that, we can then write the following LINQ expression using the `Adjustments` interface:

```csharp
Word.Shape shape = Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 100, 100);

Word.Adjustments adjustments = shape.Adjustments;

var numbers =
    from adjustment in adjustments.ToEnumerable()
    select adjustment.ToString();

foreach (var number in numbers)
{
    MessageBox.Show(number);
}
```

To sum up, there are three very different ways to use an Office collection in a LINQ expression depending on how each collection interface was defined. If your Office application uses many different collections, that can mean several different coding conventions across your code. Yuck! Not to mention having to remember which collection requires which enumeration method. This is where the Office Interop API Extensions comes in. The Office Interop API Extensions is one of the VSTO Power Tools and simplifies the use of the Office object model. Specifically, it exposes a consistent `Items()` extension method on many collection interfaces* which return the appropriate `IEnumerable<T>`. You need not remember to use `Cast<T>()` for one collection or create your own enumerator for another. Just call `Items()` and be on your way! Using our previous examples, retrieving an `IEnumerable<T>` with the Office Interop API Extensions is as simple as:

```csharp
docs.Items() // Returns IEnumerable<Word.Document>.

windows.Items() // Returns IEnumerable<Excel.Window>.

adjustments.Items() // Returns IEnumerable<float>.
```

The VSTO Power Tools are expected for release in the very near future. Keep an eye out for them and use the Office Interop API Extensions to enable LINQ in your Office applications!

*Not all of the collections are extended in this initial release. The bulk of the extensions are in the Word and Excel object models, with a key set of collections extended across the rest of the Office suite.

{% include_relative msdn-notice.md %}
