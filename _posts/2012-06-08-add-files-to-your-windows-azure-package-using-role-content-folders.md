---
layout: posts
title: "Add Files to your Windows Azure Package using Role Content Folders"
date: 2012-06-08
---
Windows Azure applications often need to package and deploy additional content. This could be advanced configuration files such as the diagnostics.wadcfg for the Diagnostics plugin. It could also be scripts and binaries that modify the virtual machine to which a role is deployed, such as installing additional runtime components. In any case, prior to the June 2012 (1.7) release of the Windows Azure Tools for Visual Studio, there were few direct ways to get this content into a package and deployed to Windows Azure.

One of the most common ways to package and deploy additional content to Windows Azure was to add the content to one of the role projects and set its Build Action to Content. On build (e.g. when creating a package or publishing to Windows Azure), the content would be copied to the role project’s output directory and subsequently picked up during packaging. There are some downsides to this approach, however. First, depending on the type of role (i.e. web or worker), the content could end up deployed to different locations, either in the role’s `APPROOT` directory or a bin subdirectory. Any component which uses this content on the virtual machine must account for that inconsistency. Second, the approach tends to pollute role projects with Azure-specific files when they might otherwise be platform agnostic. (For example, a web application that can function on-premise as well as a web role in Windows Azure.) In other words, they simply do not belong in the role projects.

The June 2012 (1.7) release introduces a new feature of Windows Azure projects: Role Content Folders. Role Content Folders are nodes within Solution Explorer to which you can add new or existing files and folders. These files will then be automatically added to the package and deployed to a consistent location on the virtual machine. The role node itself is a Role Content Folder; it represents the role’s root directory and files added to that node will be deployed to the `APPROOT` directory. Files added to folders added to the role node will be packaged and deployed in the same relative hierarchy.

To add a file, simply select the role node, right-click, and choose on of the commands from the Add submenu.

[![](/assets/posts/4403.AddNewRoleContentItem_thumb_2AD61D7B.png)](/assets/posts/6076.AddNewRoleContentItem_3D1EE43D.png)

> Note: the only item templates we include are for basic text and XML files.

[![](/assets/posts/5086.AddNewItem_thumb_1C97A48B.png)](/assets/posts/1602.AddNewItem_51A433BB.png)

Once added to a Role Content Folder, files and folders operate like any other you might find in a standard Visual Studio project. They can be added, removed, and renamed, and are also source code controlled. In this example, I’ve added two files to the web role, one directly to the web role node (i.e. `APPROOT`) and another in a subfolder.

[![](/assets/posts/5327.RoleContentHierarchy_thumb_6A33D10B.png)](/assets/posts/2671.RoleContentHierarchy_5C618B10.png)

The following two screenshots are of the application folder for one of the web role instances. In them you can see the two files deployed to the same relative hierarchy as in the Windows Azure project. That’s all it takes to package arbitrary role content!

[![](/assets/posts/4403.AppRootFile_thumb_707AA799.png)](/assets/posts/7043.AppRootFile_1101E74C.png)

[![](/assets/posts/1777.AppRootSubFolder_thumb_292551A7.png)](/assets/posts/7140.AppRootSubFolder_30448E1F.png)

You might be wondering how this all works. This new feature is actually based on an existing capability of the Windows Azure SDK. Since version 1.5 the Windows Azure SDK has supported the [Contents element](http://msdn.microsoft.com/en-us/library/windowsazure/gg557553#Contents) in the service definition that indicates additional folders to include in the package and deploy to the virtual machine. The Tools simply did not have any tooling around this capability…until now. During packaging, the Tools will identify project content within Role Content Folders and inject the appropriate elements into the service definition so that CSPack will insert that content into the package.

> Note: be careful of how you name content in a Role Content Folder. Because this content is deployed to the role’s `APPROOT` directory—the same directory in which the role is deployed—there is the possibility of file name collisions. Be sure to use filenames in your Azure project that are unlikely to conflict with files belonging to the role project itself.