---
layout: post
title: "MissingManifestResourceException when using Portable Class Libraries within WinRT"
date: 2014-11-19
---
p>Our team recently ran into a strange issue. &nbsp;Our Windows Phone application targets Silverlight 8.1 but we have several WinRT 8.1 based background tasks. &nbsp;Recently, we refactored our codebase into several Portable Class Libraries (PCLs) in order to better share code between the application and background tasks. &nbsp;Since then, however, we've started seeing <span style="font-family: 'courier new', courier;">MissingManifestResourceExceptions</span> thrown from the background tasks from the <span style="font-family: 'courier new', courier;">ResourceManager.GetString(String, CultureInfo)</span> method. &nbsp;The odd thing was that this occurred <em>only</em> in the background tasks, <em>only</em> when running on a real phone, and <em>only</em> in the application's Release configuration. &nbsp;The exceptions were not seen in the Silverlight application, or in the Windows Phone emulator, or when using the application's Debug configuration.</p>
<p>We originally thought the issue must be due to some overlooked error in our build or packaging processes, but all of our analysis indicated that the there were no significant differences in the built assemblies nor in the packages generated for Debug and Release configurations; the issue must be in the runtime. &nbsp;After a lot of digging around, we finally found an active Microsoft Connect report for the issue:</p>
<p><a title="https://connect.microsoft.com/VisualStudio/feedback/details/991028/issue-using-resx-files-on-winrt-apps-windows-phone-and-windows" href="https://connect.microsoft.com/VisualStudio/feedback/details/991028/issue-using-resx-files-on-winrt-apps-windows-phone-and-windows" target="_blank">https://connect.microsoft.com/VisualStudio/feedback/details/991028/issue-using-resx-files-on-winrt-apps-windows-phone-and-windows</a></p>
<p>That others were running into the same problem, under the same circumstances, seemed to confirm our conclusion. &nbsp;However, we still needed a workaround. &nbsp;Our options were (1) to avoid using PCLs (that use their own resources) from our background task or &nbsp;(2) redirect use of the <span style="font-family: 'courier new', courier;">ResourceManager</span> within the PCL to the WinRT <span style="font-family: 'courier new', courier;">ResourceLoader</span>. &nbsp;Option #1 would mean undoing much of our refactoring work and obviously was a non-starter.</p>
<p>Option #2 works on the assumption that, in a WinRT library, all of the resources of the referenced PCLs are extracted and placed in the package's <span style="font-family: 'courier new', courier;">resources.pri</span> file <em>in addition to</em> being embedded within the PCL assembly itself. &nbsp;That means you can get to the duplicate resources using the WinRT <span style="font-family: 'courier new', courier;">ResourceLoader</span> even if you cannot using the .NET <span style="font-family: 'courier new', courier;">ResourceManager</span>. &nbsp;Because we need to dynamically switch between using the <span style="font-family: 'courier new', courier;">ResourceManager</span> and <span style="font-family: 'courier new', courier;">ResourceLoader</span> when in Silverlight and WinRT, respectively, we need some sort of dependency injection (DI) scheme.</p>
<p>In a typical .NET library, resources are added to a Resources file (<span style="font-family: 'courier new', courier;">.resx</span>). &nbsp;The file is then used to embed the resources into the built assembly. &nbsp;The file is also used to generate a class that exposes each resource as a strongly-typed property, making it easy to consume resources within the library. &nbsp;This generated class uses a <span style="font-family: 'courier new', courier;">ResourceManager</span> instance to retrieve the resource mapped to each property. &nbsp;Redirection means either (1) not using the generated class at all and using DI scheme directly, (2) alter the generation of the class to use the DI scheme, or (3) try to alter the <em>behavior</em> of the class to conform to the DI scheme.</p>
<p>Option #1 was least desirable as it meant a lot of churn within the PCLs. &nbsp;Option #2 was not particularly desirable either as it was a significant amount of work to write and test a generator that matched the built-in one except for this one minor change. &nbsp;The question was how to implement option #3.</p>
<p>Looking at the generated class for a Resources file, each contains a static <span style="font-family: 'courier new', courier;">ResourceManager</span> field named <span style="font-family: 'courier new', courier;">resourceMan</span>. &nbsp;This field (if null)&nbsp;is set on the first retrieval of a resource via one of the generated properties. &nbsp;The <span style="font-family: 'courier new', courier;">ResourceManager.GetString(String, CultureInfo)</span> also happens to be a virtual method, which means we cancreate a derived <span style="font-family: 'courier new', courier;">ResourceManager</span> that retrieves resources from, say, a <span style="font-family: 'courier new', courier;">ResourceLoader</span>. &nbsp;The trick then, is to initialize that static <span style="font-family: 'courier new', courier;">ResourceManager</span> field to an instance of our derived <span style="font-family: 'courier new', courier;">ResourceManager</span>. &nbsp;That turns out to be a simple act of reflection as shown in the following sample:&nbsp;</p>
<p class="scroll"><span style="font-family: 'courier new', courier;"><code class="csharp">public class WindowsRuntimeResourceManager : ResourceManager</code></span></p>
<div><span style="font-family: 'courier new', courier;">{</span></div>
<p class="scroll"><span style="font-family: 'courier new', courier;">&nbsp; private readonly ResourceLoader resourceLoader;<br /> <br />&nbsp; private WindowsRuntimeResourceManager(string baseName, Assembly assembly) : base(baseName, assembly)<br />&nbsp; {<br />&nbsp; &nbsp; this.resourceLoader = ResourceLoader.GetForViewIndependentUse(baseName);<br />&nbsp; }<br /> <br />&nbsp; public static void InjectIntoResxGeneratedApplicationResourcesClass(Type resxGeneratedApplicationResourcesClass)<br />&nbsp; {<br />&nbsp; &nbsp; resxGeneratedApplicationResourcesClass.GetRuntimeFields()<br />&nbsp; &nbsp; &nbsp; .First(m =&gt; m.Name == "resourceMan")<br />&nbsp; &nbsp; &nbsp; .SetValue(null, new WindowsRuntimeResourceManager(resxGeneratedApplicationResourcesClass.FullName, resxGeneratedApplicationResourcesClass.GetTypeInfo().Assembly));<br />&nbsp; }<br /> <br />&nbsp; public override string GetString(string name, CultureInfo culture)<br />&nbsp; {<br />&nbsp; &nbsp; return this.resourceLoader.GetString(name);<br />&nbsp; }<br /> }</span></p>
<p>In the WinRT component, we then call the static injection method <em>for each</em> generated Resources class within the PCL: &nbsp; &nbsp;&nbsp;</p>
<p class="scroll"><span style="font-family: 'courier new', courier;"><code class="csharp">&nbsp; WindowsRuntimeResourceManager.InjectIntoResxGeneratedApplicationResourcesClass(typeof(PortableLibrary.Resources.AppResources));</code></span></p>
<p>With that, our WinRT background tasks and referenced PCLs were able to access all of their resources without issue and we did not have to significantly refactor any code.</p>
<p>Caveats:</p>
<ul>
<li>Using reflection to set private fields for types you do not own completely (e.g. are generated) is highly fragile and could break with the next version of the Resources class generator. &nbsp;Do this only if you have no other choice and ensure you have checks in place to detect such breaks.</li>
<li>In our case we owned all of the PCLs referenced by our WinRT component. &nbsp;I cannot say whether this workaround will work for all (e.g. third-party) PCLs, and may not work for PCLs which have internal resource types that cannot be as easily injected.</li>
</ul>