---
layout: posts
title: "Banish Missing.Value with the Office Interop API Extensions"
date: 2008-02-15
tags: [msdn]
---
I like VSTO. I like C#. What I don't like is having to write VSTO code in C# like:

```csharp
object fileName = "Test.docx";
object missing = System.Reflection.Missing.Value;

doc.SaveAs(ref fileName, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
```

This code creates a copy of a Word document, but it's not very elegant. The `SaveAs()` method accepts a long list of arguments to tweak the way the file is copied. For a simple copy, however, only the first argument--the new file name--is needed. So why am I stuck passing this "missing" value for each and every omitted argument? The definition of the SaveAs method on the Word Document COM interface defines each argument as optional (see below).  This would seem to allow callers to omit the irrelevant arguments. VB developers wouldn't give this another thought.

```csharp
HRESULT SaveAs(
    [in, optional] VARIANT* FileName,
    [in, optional] VARIANT* FileFormat,
    [in, optional] VARIANT* LockComments,
    [in, optional] VARIANT* Password,
    [in, optional] VARIANT* AddToRecentFiles,
    [in, optional] VARIANT* WritePassword,
    [in, optional] VARIANT* ReadOnlyRecommended,
    [in, optional] VARIANT* EmbedTrueTypeFonts,
    [in, optional] VARIANT* SaveNativePictureFormat,
    [in, optional] VARIANT* SaveFormsData,
    [in, optional] VARIANT* SaveAsAOCELetter,
    [in, optional] VARIANT* Encoding,
    [in, optional] VARIANT* InsertLineBreaks,
    [in, optional] VARIANT* AllowSubstitution,
    [in, optional] VARIANT* LineEnding,
    [in, optional] VARIANT* AddBiDiMask);
```

Unfortunately, the C# language and compiler do not comprehend optional arguments. What's worse, unlike the rest of the Office object model, Word interfaces use `VARIANT*` instead of `VARIANT`. That is, they are passed by reference rather than by value. This means that, not only does the C# developer have to pass a value for each and every argument, he or she must do so by reference. That means creating an extra object on the stack and then passing it to the method using the `ref` keyword. How tedious! And it gets even worse; because the values are all passed using objects, we've now lost all of the compile-time advantages of the strongly-typed C# language. How easy would it be to accidentally swap the order of the arguments?

In a perfect world, a simple copy of the document could be created like this:

```csharp
doc.SaveAs("Test.docx");
```

The method would take a strongly-typed string argument by value. And if I wanted to change the format of the newly-saved document, I could write the following:

```csharp
doc.SaveAs("Test.html", WdSaveFormat.wdFormatHTML);
```

This method would take the same string argument and another strongly-typed format specifier. We could imagine a series of method overloads that incrementally add to the argument list. But what if I need to specify a disjoint set of arguments? Why can't I simply write:

```csharp
doc.SaveAs(new DocumentSaveAsArgs
{
    FileName = "Test.docx",
    AddBiDiMarks = false
});
```

The method would take an instance of an arguments class where I could set--using the fancy new C# 3.0 object initializer syntax--only the properties necessary. Furthermore, all the properties would be strongly-typed so that any accidentally swapped arguments could be caught at compile-time and not after the application has been deployed to thousands of clients.

Am I just dreaming? Must I be content with typing `ref` over and over again for as long as I develop Office applications? Must I leave my beloved C# for the seductive VB? The answer is a resounding NO! The VSTO Power Tools announced at this week's Office Developer Conference are expected to be released in the very near future. One of those tools is the Office Interop API Extensions, a set of libraries that extend the Office object model and provide a more elegant and consistent API for the C# developer. The three examples above are all possible using the Word extensions shipped as part of this tool. Furthermore, many other interfaces* from across the Office object model have been extended in a similar manner in order to make the lives of C# developers easier. Keep an eye out for this tool and use it to banish `Missing.Value` (and its close cousin `Type.Missing`) from your C# Office applications.

*In this initial release, most of the extension work was focused on Word and Excel. The Outlook extensions had an entirely different focus which I'll discuss in a later post.

{% include_relative msdn-notice.md %}
