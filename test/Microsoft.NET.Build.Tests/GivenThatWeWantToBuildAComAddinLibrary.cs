// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.

using Newtonsoft.Json.Linq;

namespace Microsoft.NET.Build.Tests
{
    public class GivenThatWeWantToBuildAComAddinLibrary : SdkTest
    {
        public GivenThatWeWantToBuildAComAddinLibrary(ITestOutputHelper log) : base(log)
        {
        }

        [WindowsOnlyFact]
        public void It_copies_the_comhost_to_the_output_directory()
        {
            var testAsset = _testAssetsManager
                .CopyTestAsset("ComAddin")
                .WithSource();

            var buildCommand = new BuildCommand(testAsset);
            buildCommand
                .Execute()
                .Should()
                .Pass();

            var outputDirectory = buildCommand.GetOutputDirectory(ToolsetInfo.CurrentTargetFramework);

            outputDirectory.Should().OnlyHaveFiles(new[] {
                "ComAddin.dll",
                "ComAddin.pdb",
                "ComAddin.deps.json",
                "ComAddin.comhost.win-x86.dll",
                "ComAddin.comhost.win-x64.dll",
                "ComAddin.comhost.win-arm64.dll",
                "ComAddin.runtimeconfig.json"
            });

            string runtimeConfigFile = Path.Combine(outputDirectory.FullName, "ComAddin.runtimeconfig.json");
            string runtimeConfigContents = File.ReadAllText(runtimeConfigFile);
            JObject runtimeConfig = JObject.Parse(runtimeConfigContents);
            runtimeConfig["runtimeOptions"]["rollForward"].Value<string>()
                .Should().Be("LatestMinor");
        }
    }
}
