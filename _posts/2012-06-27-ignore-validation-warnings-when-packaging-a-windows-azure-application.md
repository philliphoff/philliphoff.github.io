---
layout: posts
title: "Ignore Validation Warnings when Packaging a Windows Azure Application"
date: 2012-06-27
---
When a package fails to deploy to Windows Azure (or deploys but its roles fail to start properly) it can be difficult to determine what went wrong. In many cases, these failures often happen before the diagnostics agent has had a chance to startup and help identify the issue. A lot of these failures are due to simple but easily forgotten tasks, such as not adding a referenced assembly to the package or not updating your connection strings to point to Windows Azure storage instead of the storage emulator. The Windows Azure Tools for Visual Studio tries to help developers catch many of these issues before they spend time waiting for their package to deploy and the application to start (and then fail). When such issues are encountered during packaging the build process generates validation warnings or errors. The developer then has immediate, specific feedback on what needs to be done before deploying again to Windows Azure.

Validation performed by the Tools include:

 - Verify that the projects corresponding to each role are built against versions of the .NET runtime that exist on the virtual machines.
 - Verify that the projects reference only assemblies that exist on the virtual machines (i.e. the .NET Framework or Windows Azure client libraries) or have been specifically included in the package (i.e. Copy Local is True).
 - Verify that connection strings for diagnostics or caching in the service configuration (.cscfg) do not point to the storage emulator (which will likely cause the corresponding agents to fail, and thus the role to fail).

In some cases, however, these warnings may not apply. For example, suppose you are generating a package as part of a continuous build that will then be deployed to the development fabric and undergo a series of automated tests. If you were to generate a package (e.g. by calling the Publish target), you may see the following warnings:

```
WebRole1(0,0): warning WAT170: The configuration setting 'Microsoft.WindowsAzure.Plugins.Diagnostics.ConnectionString' is set up to use the local storage emulator for role 'WebRole1' in configuration file 'ServiceConfiguration.Cloud.cscfg'. To access Windows Azure storage services, you must provide a valid Windows Azure storage connection string.
WebRole1(0,0): warning WAT230: The connection string 'DefaultConnection' is using a local database '(LocalDb)\v11.0' in project 'WebRole1'. This connection string will not work when you run this application in Windows Azure. To access a different database, you should update the connection string in the web.config file.
```

The warnings indicate that there are connection strings pointing to the storage emulator as well as a local database and these will likely cause a deployment to Windows Azure to fail. In this case, however, we know that the package is being deployed to the development fabric so those connection strings are, in fact, perfectly valid. In the interests of producing "clean" builds, it would be nice to ignore these warnings in this particular case.

To ignore validation warnings when packaging a Windows Azure application, set the `IgnoreValidationIssueCodes` MSBuild property in your application's project file. For example, to ignore the two warnings encountered above set the property after the import of the WindowsAzure targets file:

```
.
.
.
<Import Project="$(CloudExtensionsDir)Microsoft.WindowsAzure.targets" />
<PropertyGroup>
  <IgnoreValidationIssueCodes>WAT170;WAT230</IgnoreValidationIssueCodes>
</PropertyGroup>
. 
.
.
```

Now these warnings will not be generated when packaging.