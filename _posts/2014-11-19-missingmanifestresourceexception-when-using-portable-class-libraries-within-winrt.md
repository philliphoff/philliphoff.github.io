---
layout: posts
title: "MissingManifestResourceException when using Portable Class Libraries within WinRT"
date: 2014-11-19
---
Our team recently ran into a strange issue. Our Windows Phone application targets Silverlight 8.1 but we have several WinRT 8.1 based background tasks. Recently, we refactored our codebase into several Portable Class Libraries (PCLs) in order to better share code between the application and background tasks. Since then, however, we've started seeing `MissingManifestResourceExceptions` thrown from the background tasks from the `ResourceManager.GetString(String, CultureInfo)` method. The odd thing was that this occurred *only* in the background tasks, *only* when running on a real phone, and *only* in the application's `Release` configuration. The exceptions were not seen in the Silverlight application, or in the Windows Phone emulator, or when using the application's Debug configuration.

We originally thought the issue must be due to some overlooked error in our build or packaging processes, but all of our analysis indicated that the there were no significant differences in the built assemblies nor in the packages generated for `Debug` and `Release` configurations; the issue must be in the runtime. After a lot of digging around, we finally found an active Microsoft Connect report for the issue:

[https://connect.microsoft.com/VisualStudio/feedback/details/991028/issue-using-resx-files-on-winrt-apps-windows-phone-and-windows](https://connect.microsoft.com/VisualStudio/feedback/details/991028/issue-using-resx-files-on-winrt-apps-windows-phone-and-windows)

That others were running into the same problem, under the same circumstances, seemed to confirm our conclusion. However, we still needed a workaround. Our options were (1) to avoid using PCLs (that use their own resources) from our background task or (2) redirect use of the `ResourceManager` within the PCL to the WinRT `ResourceLoader`.

Option #1 would mean undoing much of our refactoring work and obviously was a non-starter.

Option #2 works on the assumption that, in a WinRT library, all of the resources of the referenced PCLs are extracted and placed in the package's `resources.pri` file *in addition to* being embedded within the PCL assembly itself. That means you can get to the duplicate resources using the WinRT `ResourceLoader` even if you cannot using the .NET `ResourceManager`. Because we need to dynamically switch between using the `ResourceManager` and `ResourceLoader` when in Silverlight and WinRT, respectively, we need some sort of dependency injection (DI) scheme.

In a typical .NET library, resources are added to a Resources file (`.resx`). The file is then used to embed the resources into the built assembly. The file is also used to generate a class that exposes each resource as a strongly-typed property, making it easy to consume resources within the library. This generated class uses a `ResourceManager` instance to retrieve the resource mapped to each property. Redirection means either (1) not using the generated class at all and using DI scheme directly, (2) alter the generation of the class to use the DI scheme, or (3) try to alter the *behavior* of the class to conform to the DI scheme.

Option #1 was least desirable as it meant a lot of churn within the PCLs. Option #2 was not particularly desirable either as it was a significant amount of work to write and test a generator that matched the built-in one except for this one minor change. The question was how to implement option #3.

Looking at the generated class for a Resources file, each contains a static `ResourceManager` field named `resourceMan`. This field (if null) is set on the first retrieval of a resource via one of the generated properties. The `ResourceManager.GetString(String, CultureInfo)` also happens to be a virtual method, which means we cancreate a derived `ResourceManager` that retrieves resources from, say, a `ResourceLoader`. The trick then, is to initialize that static `ResourceManager` field to an instance of our derived `ResourceManager`. That turns out to be a simple act of reflection as shown in the following sample:

```csharp
public class WindowsRuntimeResourceManager : ResourceManager
{
    private readonly ResourceLoader resourceLoader;

    private WindowsRuntimeResourceManager(string baseName, Assembly assembly) : base(baseName, assembly)
    {
        this.resourceLoader = ResourceLoader.GetForViewIndependentUse(baseName);
    }

    public static void InjectIntoResxGeneratedApplicationResourcesClass(Type resxGeneratedApplicationResourcesClass)
    {
        resxGeneratedApplicationResourcesClass
            .GetRuntimeFields()
            .First(m => m.Name == "resourceMan")
            .SetValue(null, new WindowsRuntimeResourceManager(resxGeneratedApplicationResourcesClass.FullName, resxGeneratedApplicationResourcesClass.GetTypeInfo().Assembly));
    }

    public override string GetString(string name, CultureInfo culture)
    {
        return this.resourceLoader.GetString(name);
    }
}
```

In the WinRT component, we then call the static injection method *for each* generated Resources class within the PCL:

```csharp
WindowsRuntimeResourceManager.InjectIntoResxGeneratedApplicationResourcesClass(typeof(PortableLibrary.Resources.AppResources));
```

With that, our WinRT background tasks and referenced PCLs were able to access all of their resources without issue and we did not have to significantly refactor any code.

Caveats:

 - Using reflection to set private fields for types you do not own completely (e.g. are generated) is highly fragile and could break with the next version of the `Resources` class generator. Do this only if you have no other choice and ensure you have checks in place to detect such breaks.

 - In our case we owned all of the PCLs referenced by our WinRT component. I cannot say whether this workaround will work for all (e.g. third-party) PCLs, and may not work for PCLs which have internal resource types that cannot be as easily injected.
