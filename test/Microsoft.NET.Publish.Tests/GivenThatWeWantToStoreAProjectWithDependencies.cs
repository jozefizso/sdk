// Copyright (c) .NET Foundation and contributors. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using FluentAssertions;
using Microsoft.NET.TestFramework;
using Microsoft.NET.TestFramework.Assertions;
using Microsoft.NET.TestFramework.Commands;
using NuGet.Packaging.Core;
using NuGet.Versioning;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Xunit;
using Xunit.Abstractions;

namespace Microsoft.NET.Publish.Tests
{
    public class GivenThatWeWantToStoreAProjectWithDependencies : SdkTest
    {
        private static readonly string _libPrefix = FileConstants.DynamicLibPrefix;
        private static string _runtimeRid;
        private static string _testArch;
        private static string _tfm = "netcoreapp2.0";

        static GivenThatWeWantToStoreAProjectWithDependencies()
        {
            var rid = Microsoft.DotNet.PlatformAbstractions.RuntimeEnvironment.GetRuntimeIdentifier();
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                _testArch = rid.Substring(rid.LastIndexOf("-") + 1);
                _runtimeRid = "win7-" + _testArch;
            }
            else
            {

                if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                {
                    // microsoft.netcore.coredistools only has assets for osx.10.10
                    _runtimeRid = "osx.10.10-x64";
                }
                else if (rid.Contains("ubuntu"))
                {
                    // microsoft.netcore.coredistools only has assets for ubuntu.14.04-x64
                    _runtimeRid = "ubuntu.14.04-x64";
                }
                else
                {
                    _runtimeRid = rid;
                }
            }
        }

        public GivenThatWeWantToStoreAProjectWithDependencies(ITestOutputHelper log) : base(log)
        {
        }

        [CoreMSBuildOnlyFact]
        public void compose_dependencies()
        {
            TestAsset simpleDependenciesAsset = _testAssetsManager
                .CopyTestAsset("TargetManifests")
                .WithSource();

            var storeCommand = new ComposeStoreCommand(Log, simpleDependenciesAsset.TestRoot, "FluentAssertion.xml");

            var OutputFolder = Path.Combine(simpleDependenciesAsset.TestRoot, "outdir");
            var WorkingDir = Path.Combine(simpleDependenciesAsset.TestRoot, "w");

            storeCommand
                .Execute($"/p:RuntimeIdentifier={_runtimeRid}", $"/p:TargetFramework={_tfm}", $"/p:ComposeDir={OutputFolder}", $"/p:ComposeWorkingDir={WorkingDir}", "/p:DoNotDecorateComposeDir=true", "/p:PreserveComposeWorkingDir=true", "/p:CreateProfilingSymbols=false")
                .Should()
                .Pass();
            DirectoryInfo storeDirectory = new DirectoryInfo(OutputFolder);

            List<string> files_on_disk = new List<string> {
               "artifact.xml",
               @"newtonsoft.json/9.0.1/lib/netstandard1.0/Newtonsoft.Json.dll",
               @"fluentassertions/4.12.0/lib/netstandard1.3/FluentAssertions.Core.dll",
               @"fluentassertions/4.12.0/lib/netstandard1.3/FluentAssertions.dll",
               @"fluentassertions.json/4.12.0/lib/netstandard1.3/FluentAssertions.Json.dll"
               };

           
            storeDirectory.Should().HaveFiles(files_on_disk);
        }

        [CoreMSBuildOnlyFact]
        public void compose_dependencies_noopt()
        {
            TestAsset simpleDependenciesAsset = _testAssetsManager
                .CopyTestAsset("TargetManifests")
                .WithSource();

            var storeCommand = new ComposeStoreCommand(Log, simpleDependenciesAsset.TestRoot, "FluentAssertion.xml");

            var OutputFolder = Path.Combine(simpleDependenciesAsset.TestRoot, "outdir");
            var WorkingDir = Path.Combine(simpleDependenciesAsset.TestRoot, "w");

            storeCommand
                .Execute($"/p:RuntimeIdentifier={_runtimeRid}", $"/p:TargetFramework={_tfm}", $"/p:ComposeDir={OutputFolder}", "/p:SkipOptimization=true", $"/p:ComposeWorkingDir={WorkingDir}", "/p:DoNotDecorateComposeDir=true", "/p:PreserveComposeWorkingDir=true", "/p:CreateProfilingSymbols=false")
                .Should()
                .Pass();
            DirectoryInfo storeDirectory = new DirectoryInfo(OutputFolder);

            List<string> files_on_disk = new List<string> {
               "artifact.xml",
               @"newtonsoft.json/9.0.1/lib/netstandard1.0/Newtonsoft.Json.dll",
               @"fluentassertions/4.12.0/lib/netstandard1.3/FluentAssertions.Core.dll",
               @"fluentassertions/4.12.0/lib/netstandard1.3/FluentAssertions.dll",
               @"fluentassertions.json/4.12.0/lib/netstandard1.3/FluentAssertions.Json.dll"
               };

           
            storeDirectory.Should().HaveFiles(files_on_disk);

        }

        [CoreMSBuildOnlyFact]
        public void store_nativeonlyassets()
        {
            TestAsset simpleDependenciesAsset = _testAssetsManager
                .CopyTestAsset("UnmanagedStore")
                .WithSource();

            var storeCommand = new ComposeStoreCommand(Log, simpleDependenciesAsset.TestRoot);

            var OutputFolder = Path.Combine(simpleDependenciesAsset.TestRoot, "outdir");
            var WorkingDir = Path.Combine(simpleDependenciesAsset.TestRoot, "w");
            storeCommand
                .Execute($"/p:RuntimeIdentifier={_runtimeRid}", $"/p:TargetFramework={_tfm}", $"/p:ComposeWorkingDir={WorkingDir}", $"/p:ComposeDir={OutputFolder}", $"/p:DoNotDecorateComposeDir=true")
                .Should()
                .Pass();

            DirectoryInfo storeDirectory = new DirectoryInfo(OutputFolder);

            List<string> files_on_disk = new List<string> {
               "artifact.xml",
               $"runtime.{_runtimeRid}.microsoft.netcore.coredistools/1.0.1-prerelease-00001/runtimes/{_runtimeRid}/native/{_libPrefix}coredistools{FileConstants.DynamicLibSuffix}",
               $"runtime.{_runtimeRid}.microsoft.netcore.coredistools/1.0.1-prerelease-00001/runtimes/{_runtimeRid}/native/coredistools.h"
               };

            storeDirectory.Should().OnlyHaveFiles(files_on_disk);
        }

        [CoreMSBuildOnlyFact]
        public void compose_multifile()
        {
            TestAsset simpleDependenciesAsset = _testAssetsManager
                .CopyTestAsset("TargetManifests", "multifile")
                .WithSource();

            var storeCommand =  new ComposeStoreCommand(Log, simpleDependenciesAsset.TestRoot, "NewtonsoftFilterProfile.xml");

            var OutputFolder = Path.Combine(simpleDependenciesAsset.TestRoot, "o");
            var WorkingDir = Path.Combine(simpleDependenciesAsset.TestRoot, "w");
            var additonalproj1 = Path.Combine(simpleDependenciesAsset.TestRoot, "NewtonsoftMultipleVersions.xml");
            var additonalproj2 = Path.Combine(simpleDependenciesAsset.TestRoot, "FluentAssertion.xml");

            storeCommand
                .Execute($"/p:RuntimeIdentifier={_runtimeRid}", $"/p:TargetFramework={_tfm}", $"/p:Additionalprojects={additonalproj1}%3b{additonalproj2}", $"/p:ComposeDir={OutputFolder}", $"/p:ComposeWorkingDir={WorkingDir}", "/p:DoNotDecorateComposeDir=true", "/p:CreateProfilingSymbols=false")
                .Should()
                .Pass();
            DirectoryInfo storeDirectory = new DirectoryInfo(OutputFolder);

            List<string> files_on_disk = new List<string> {
               "artifact.xml",
               @"newtonsoft.json/9.0.2-beta2/lib/netstandard1.1/Newtonsoft.Json.dll",
               @"newtonsoft.json/9.0.1/lib/netstandard1.0/Newtonsoft.Json.dll",
               @"fluentassertions.json/4.12.0/lib/netstandard1.3/FluentAssertions.Json.dll"
               };

            storeDirectory.Should().HaveFiles(files_on_disk);

            var knownpackage = new HashSet<PackageIdentity>();

            knownpackage.Add(new PackageIdentity("Newtonsoft.Json", NuGetVersion.Parse("9.0.1")));
            knownpackage.Add(new PackageIdentity("Newtonsoft.Json", NuGetVersion.Parse("9.0.2-beta2")));
            knownpackage.Add(new PackageIdentity("FluentAssertions.Json", NuGetVersion.Parse("4.12.0")));

            var artifact = Path.Combine(OutputFolder, "artifact.xml");
            var packagescomposed = ParseStoreArtifacts(artifact);

            packagescomposed.Count.Should().BeGreaterThan(0);

            foreach (var pkg in knownpackage)
            {
                packagescomposed.Should().Contain(elem => elem.Equals(pkg), "package {0}, version {1} was not expected to be stored", pkg.Id, pkg.Version);
            }
        }

        [CoreMSBuildOnlyFact]
        public void It_uses_star_versions_correctly()
        {
            TestAsset targetManifestsAsset = _testAssetsManager
                .CopyTestAsset("TargetManifests")
                .WithSource();

            var outputFolder = Path.Combine(targetManifestsAsset.TestRoot, "o");
            var workingDir = Path.Combine(targetManifestsAsset.TestRoot, "w");

            new ComposeStoreCommand(Log, targetManifestsAsset.TestRoot, "StarVersion.xml")
                .Execute($"/p:RuntimeIdentifier={_runtimeRid}", $"/p:TargetFramework={_tfm}", $"/p:ComposeDir={outputFolder}", $"/p:ComposeWorkingDir={workingDir}", "/p:DoNotDecorateComposeDir=true", "/p:CreateProfilingSymbols=false")
                .Should()
                .Pass();

            var artifactFile = Path.Combine(outputFolder, "artifact.xml");
            var storeArtifacts = ParseStoreArtifacts(artifactFile);

            var nugetPackage = storeArtifacts.Single(p => string.Equals(p.Id, "NuGet.Common", StringComparison.OrdinalIgnoreCase));

            // nuget.org/packages/NuGet.Common currently contains:
            // 4.0.0
            // 4.0.0-rtm-2283
            // 4.0.0-rtm-2265
            // 4.0.0-rc3
            // 4.0.0-rc2
            // 4.0.0-rc-2048
            //
            // and the StarVersion.xml uses Version="4.0.0-*", 
            // so we expect a version greater than 4.0.0-rc2, since there is
            // a higher version on the feed that meets the criteria
            nugetPackage.Version.Should().BeGreaterThan(NuGetVersion.Parse("4.0.0-rc2"));
        }

        [CoreMSBuildOnlyFact]
        public void It_creates_profiling_symbols()
        {
            TestAsset targetManifestsAsset = _testAssetsManager
                .CopyTestAsset("TargetManifests")
                .WithSource();

            var outputFolder = Path.Combine(targetManifestsAsset.TestRoot, "o");
            var workingDir = Path.Combine(targetManifestsAsset.TestRoot, "w");

            var composeStore = new ComposeStoreCommand(Log, targetManifestsAsset.TestRoot, "NewtonsoftFilterProfile.xml");

            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                // clear the PATH on windows to ensure creating .ni.pdbs works without 
                // being in a VS developer command prompt
                composeStore.WithEnvironmentVariable("PATH", string.Empty);
            }
            
            composeStore
                .Execute(
                    $"/p:RuntimeIdentifier={_runtimeRid}",
                    "/p:TargetFramework=netcoreapp2.0",
                    $"/p:ComposeDir={outputFolder}",
                    $"/p:ComposeWorkingDir={workingDir}",
                    "/p:DoNotDecorateComposeDir=true",
                    "/p:PreserveComposeWorkingDir=true")
                .Should()
                .Pass();

            var symbolFileExtension = RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? "ni.pdb" : ".map";
            var symbolsFolder = new DirectoryInfo(Path.Combine(outputFolder, "symbols"));

            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
                // profiling symbols are not supported on OSX
                symbolsFolder.Should().NotExist();
                return;
            }

            var newtonsoftSymbolsFolder = symbolsFolder.Sub("newtonsoft.json").Sub("9.0.1").Sub("lib").Sub("netstandard1.0");
            newtonsoftSymbolsFolder.Should().Exist();

            var newtonsoftSymbolsFiles = newtonsoftSymbolsFolder.GetFiles().ToArray();
            newtonsoftSymbolsFiles.Length.Should().Be(1);
            newtonsoftSymbolsFiles[0].Name.Should().StartWith("Newtonsoft.Json").And.EndWith(symbolFileExtension);
        }

        private static HashSet<PackageIdentity> ParseStoreArtifacts(string path)
        {
            return new HashSet<PackageIdentity>(
                from element in XDocument.Load(path).Root.Elements("Package")
                select new PackageIdentity(
                    element.Attribute("Id").Value,
                    NuGetVersion.Parse(element.Attribute("Version").Value)));
        }
    }
}
