using System.Linq;
using Nuke.Common;
using Nuke.Common.Execution;
using Nuke.Common.Git;
using Nuke.Common.IO;
using Nuke.Common.ProjectModel;
using Nuke.Common.Tooling;
using Nuke.Common.Tools.DotNet;
using Nuke.Common.Tools.GitVersion;
using Nuke.Common.Tools.MSBuild;
using Nuke.Common.Tools.Xunit;
using Nuke.Common.Utilities.Collections;
using static Nuke.Common.IO.FileSystemTasks;
using static Nuke.Common.IO.PathConstruction;
using static Nuke.Common.Tools.MSBuild.MSBuildTasks;
using static Nuke.Common.Tools.Xunit.XunitTasks;
using static Nuke.Common.Tools.DotNet.DotNetTasks;

[CheckBuildProjectConfigurations]
[UnsetVisualStudioEnvironmentVariables]
class Build : NukeBuild
{
    public static int Main() => Execute<Build>(x => x.Ci);

    [Parameter("Configuration to build - Default is 'Debug' (local) or 'Release' (server)")]
    readonly Configuration Configuration = IsLocalBuild ? Configuration.Debug : Configuration.Release;

    [Solution] readonly Solution Solution;
    [GitRepository] readonly GitRepository GitRepository;
    [GitVersion] readonly GitVersion GitVersion;

    AbsolutePath SourceDirectory => RootDirectory / "src";
    AbsolutePath ArtifactsDirectory => RootDirectory / "artifacts";

    Target Clean => _ => _
        .Before(Restore)
        .Executes(() =>
        {
            SourceDirectory.GlobDirectories("**/bin", "**/obj").ForEach(DeleteDirectory);
            EnsureCleanDirectory(ArtifactsDirectory);
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
                .SetTargetPath(SourceDirectory / "Allors.Excel.Tests" / "Allors.Excel.Tests.csproj")
                .SetTargets("Rebuild")
                .SetConfiguration(Configuration)
                .SetAssemblyVersion(GitVersion.AssemblySemVer)
                .SetFileVersion(GitVersion.AssemblySemFileVer)
                .SetInformationalVersion(GitVersion.InformationalVersion));

        });

    Target Tests => _ => _
       .DependsOn(Compile)
       .Executes(() =>
       {
           var assembly = (AbsolutePath)GlobFiles(SourceDirectory, "**/Allors.Excel.Tests.dll").First();
           var workingDirectory = assembly.Parent;

           Xunit2(v => v
                 .SetFramework("net461")
                 .SetWorkingDirectory(workingDirectory)
                 .AddTargetAssemblies(assembly)
                 .SetResultReport(Xunit2ResultFormat.Xml, ArtifactsDirectory / "tests" / "results.xml"));
       });

    Target Pack => _ => _
       .After(Tests)
       .DependsOn(Compile)
       .Executes(() =>
       {
           var projects = new[] { "Allors.Excel", "Allors.Excel.Headless", "Allors.Excel.Interop" };

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
       });

    Target CiTests => _ => _
    .DependsOn(Tests);

    Target Ci => _ => _
        .DependsOn(Pack, Tests);
}
