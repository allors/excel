using System.Linq;
using Nuke.Common;
using Nuke.Common.IO;
using Nuke.Common.ProjectModel;
using Nuke.Common.Tooling;
using Nuke.Common.Tools.Coverlet;
using Nuke.Common.Tools.MSBuild;
using Nuke.Common.Tools.DotNet;
using Nuke.Common.Tools.GitVersion;
using Nuke.Common.Tools.Xunit;
using static Nuke.Common.Tools.MSBuild.MSBuildTasks;
using static Nuke.Common.Tools.Xunit.XunitTasks;
using static Nuke.Common.Tools.DotNet.DotNetTasks;

class Build : NukeBuild
{
    public static int Main() => Execute<Build>(x => x.Default);

    [Parameter("Configuration to build - Default is 'Debug' (local) or 'Release' (server)")]
    readonly Configuration Configuration = IsLocalBuild ? Configuration.Debug : Configuration.Release;

    [Solution] readonly Solution Solution;
    [GitVersion] readonly GitVersion GitVersion;

    AbsolutePath SourceDirectory => RootDirectory / "src";
    AbsolutePath ArtifactsDirectory => RootDirectory / "artifacts";
    AbsolutePath TestsDirectory => ArtifactsDirectory / "tests";

    Target Clean => _ => _
        .Before(Restore)
        .Executes(() =>
        {
            // DotNetClean(v => v.SetProject(SourceDirectory));
            ArtifactsDirectory.CreateOrCleanDirectory();
        });

    Target Restore => _ => _
      .Executes(() =>
      {
          MSBuild(s => s
              .SetTargetPath(Solution)
              .SetTargets("Restore"));
      });

    Target Compile => _ => _
        .DependsOn(Restore)
        .Executes(() =>
        {
            MSBuild(s => s
                .SetTargetPath(SourceDirectory / "Allors.Excel.Headless.Tests" / "Allors.Excel.Headless.Tests.csproj")
                .SetTargets("Rebuild")
                .SetConfiguration(Configuration)
                .SetAssemblyVersion(GitVersion.AssemblySemVer)
                .SetFileVersion(GitVersion.AssemblySemFileVer)
                .SetInformationalVersion(GitVersion.InformationalVersion));

            MSBuild(s => s
                .SetTargetPath(SourceDirectory / "Allors.Excel.Interop.Tests" / "Allors.Excel.Interop.Tests.csproj")
                .SetTargets("Rebuild")
                .SetConfiguration(Configuration)
                .SetAssemblyVersion(GitVersion.AssemblySemVer)
                .SetFileVersion(GitVersion.AssemblySemFileVer)
                .SetInformationalVersion(GitVersion.InformationalVersion));

            MSBuild(s => s
                .SetTargetPath(SourceDirectory / "Allors.Excel.Interop.Vsto.Tests" / "Allors.Excel.Interop.Vsto.Tests.csproj")
                .SetTargets("Rebuild")
                .SetConfiguration(Configuration)
                .SetAssemblyVersion(GitVersion.AssemblySemVer)
                .SetFileVersion(GitVersion.AssemblySemFileVer)
                .SetInformationalVersion(GitVersion.InformationalVersion));
        });

    Target Test => _ => _
       .DependsOn(Compile)
       .Executes(() =>
       {
           DotNetTest(s => s
               .SetProjectFile(Solution.GetProject("Allors.Excel.Interop.Tests"))
               .SetConfiguration(Configuration)
               .EnableNoBuild()
               .EnableNoRestore()
               .AddLoggers("trx;LogFileName=interop-results-dotnet.trx")
               .EnableProcessLogOutput()
               .SetResultsDirectory(TestsDirectory)
           );

           {
               var assembly = SourceDirectory.GlobFiles("**/Allors.Excel.Interop.Tests.dll").First();

               Xunit2(v => v
                   .SetFramework("net462")
                   .AddTargetAssemblies(assembly)
                   .SetResultReport(Xunit2ResultFormat.Xml, TestsDirectory / "interop-results-framework.xml"));
           }

           {
               var assembly = SourceDirectory.GlobFiles("**/Allors.Excel.Interop.Vsto.Tests.dll").First();

               Xunit2(v => v
                   .SetFramework("net462")
                   .AddTargetAssemblies(assembly)
                   .SetResultReport(Xunit2ResultFormat.Xml, TestsDirectory / "interop-vsto-results.xml"));
           }

           {
               var assembly = SourceDirectory.GlobFiles("**/Allors.Excel.Headless.Tests.dll").First();

               Xunit2(v => v
                   .SetFramework("net462")
                   .AddTargetAssemblies(assembly)
                   .SetResultReport(Xunit2ResultFormat.Xml, TestsDirectory / "headless-results.xml"));
           }
       });

    Target Pack => _ => _
       .After(Test)
       .DependsOn(Compile)
       .Executes(() =>
       {
           var projects = new[] { "Allors.Excel", "Allors.Excel.Headless" };

           foreach (var project in projects)
           {
               DotNetPack(s => s
                .SetProject(Solution.GetProject(project))
                .SetConfiguration(Configuration)
                .EnableIncludeSource()
                .EnableIncludeSymbols()
                .SetVersion(GitVersion.NuGetVersionV2)
                .SetOutputDirectory(ArtifactsDirectory / "nuget"));
           }

           MSBuild(s => s
               .SetTargetPath(SourceDirectory / "Allors.Excel.Interop" / "Allors.Excel.Interop.csproj")
               .SetTargets("Pack")
               .SetConfiguration(Configuration)
               .SetPackageVersion(GitVersion.AssemblySemVer)
               .SetAssemblyVersion(GitVersion.AssemblySemVer)
               .SetFileVersion(GitVersion.AssemblySemFileVer)
               .SetInformationalVersion(GitVersion.InformationalVersion)
               .SetPackageOutputPath(ArtifactsDirectory / "nuget"));
       });

    Target Ci => _ => _
    .DependsOn(Test);

    Target Default => _ => _
        .DependsOn(Pack, Test);
}
