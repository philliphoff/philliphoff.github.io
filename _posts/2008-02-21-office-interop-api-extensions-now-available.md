---
layout: posts
title: "Office Interop API Extensions Now Available!"
date: 2008-02-21
tags: [msdn]
---
As announced in Andrew Whitechapel's [post](http://blogs.msdn.com/andreww/archive/2008/02/21/vsto-vsta-power-tools-v1-0.aspx), version 1.0 of the VSTO Power Tools have been released! One of those tools is the Office Interop API Extensions, a set of libraries which extend the Office object model to simplify development on the Office platform. This past week I've blogged about the capabilities of these extensions and now you have the chance to try them out yourselves! Please let us know what works, what doesn't, what could be improved, and what you would like to see in the future.

The VSTO Power Tools v1.0.0.0 is a collection of three packages found [here](http://www.microsoft.com/downloads/details.aspx?FamilyId=46B6BF86-E35D-4870-B214-4D7B72B02BF9&amp;displaylang=en). The Office Interop API Extensions is packaged as `VSTO_PTExtLibs.exe` Please note that, despite the prerequisites listed on the MSDN download page, the Office Interop API Extensions can be used with both Office 2007 **and** Office 2003, though .NET 3.5 is required in both cases. Also, these extensions do **not** have a dependency on VSTO; they can be used within any application which automates Office.

{% include_relative msdn-notice.md %}
