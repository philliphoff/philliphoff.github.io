---
layout: posts
title: "Finding DASL Property Names"
date: 2008-12-19
---
The LINQ-to-DASL provider of the Office Interop API Extensions provides a very limited set of mappings between its query types and their associated DASL properties. We didn't have the time to add them all and we didn't know which properties (besides the obvious ones like `Subject` and `DateReceived`) users would be most likely to use in their queries. Instead, we created an extensibility mechanism (the `OutlookItemPropertyAttribute`) that allowed users to add their own mappings to DASL properties. Now the problem becomes, where do we find these DASL property names?

You can find some of the property names amongst the MSDN Office and Exchange documentation, but many properties are missing and sometimes the documentation is misleading (e.g. a listed property may work when querying Exchange but not Outlook). Outlook supports another query syntax (Jet) that uses a different set of property names, making the documentation even more confusing.

The best approach that I have found is to go directly to the source. Outlook actually has a built-in DASL query builder, if only you know how to find it. Here's the trick: start with the View : Current View : Customize Current View menu item.

[![](/assets/posts/ViewMenu_thumb.jpg)](/assets/posts/ViewMenu.jpg)

This displays the Customize View Messages dialog.

[![](/assets/posts/CustomizeCurrentView_thumb.jpg)](/assets/posts/CustomizeCurrentView.jpg)

Next, click the Filter button.

[![](/assets/posts/AdvancedTab_thumb.jpg)](/assets/posts/AdvancedTab.jpg)

The filter dialog allows you to filter messages in a folder by a variety of criteria. Select the Advanced tab.

[![](/assets/posts/FieldMenu_thumb.jpg)](/assets/posts/FieldMenu_2.jpg)

Now use the Field button to select the Outlook item property you are interested in.

[![](/assets/posts/FieldConditionAndValue_thumb.jpg)](/assets/posts/FieldConditionAndValue.jpg)

You can then select the condition and a value. Click the Add to List button to add the criteria to the filter.

[![](/assets/posts/CriteriaAdded_thumb.jpg)](/assets/posts/CriteriaAdded.jpg)

Finally, select the SQL tab to see the DASL.

[![](/assets/posts/SqlTab_thumb.jpg)](/assets/posts/SqlTab.jpg)

In this case, we see that the DASL property associated with the Modified time of an appointment is `"DAV:getlastmodified"`. To cut and paste the query string elsewhere (e.g. to Notepad, so you can compare the query string to the [debug output of your LINQ-to-DASL query]({% post_url 2008-12-18-debugging-linq-to-dasl-queries %})), check the box at the bottom of the tab.

[![](/assets/posts/SelectedDasl_thumb.jpg)](/assets/posts/SelectedDasl.jpg)
