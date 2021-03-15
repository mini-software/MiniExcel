<Project DefaultTargets="Build" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <Import Project="Microsoft.CodeDom.Providers.DotNetCompilerPlatform.Extensions.props"/>

  <ItemGroup>
    <RoslynCompilerFiles Include="$(RoslynToolPath)\*">
      <Link>roslyn\%(RecursiveDir)%(Filename)%(Extension)</Link>
    </RoslynCompilerFiles>
  </ItemGroup>
  <Target Name="IncludeRoslynCompilerFilesToFilesForPackagingFromProject" BeforeTargets="PipelineCollectFilesPhase" >
    <ItemGroup>
      <FilesForPackagingFromProject Include="@(RoslynCompilerFiles)">
        <DestinationRelativePath>bin\roslyn\%(RecursiveDir)%(Filename)%(Extension)</DestinationRelativePath>
        <FromTarget>IncludeRoslynCompilerFilesToFilesForPackagingFromProject</FromTarget>
        <Category>Run</Category>
      </FilesForPackagingFromProject>
    </ItemGroup>
  </Target>
  <Target Name="LocateRoslynToolsDestinationFolder" Condition=" '$(RoslynToolsDestinationFolder)' == '' ">
      <PropertyGroup>
        <RoslynToolsDestinationFolder>$(WebProjectOutputDir)\bin\roslyn</RoslynToolsDestinationFolder>
        <RoslynToolsDestinationFolder Condition=" '$(WebProjectOutputDir)' == '' ">$(OutputPath)\roslyn</RoslynToolsDestinationFolder>
    </PropertyGroup>
  </Target>
  <Target Name="CopyRoslynCompilerFilesToOutputDirectory" AfterTargets="CopyFilesToOutputDirectory" DependsOnTargets="LocateRoslynToolsDestinationFolder">
    <Copy SourceFiles="@(RoslynCompilerFiles)" DestinationFolder="$(RoslynToolsDestinationFolder)" ContinueOnError="true" SkipUnchangedFiles="true" Retries="0" />
    <ItemGroup  Condition="'$(MSBuildLastTaskResult)' == 'True'" >
      <FileWrites Include="$(RoslynToolsDestinationFolder)\*" />
    </ItemGroup>
  </Target>
  <Target Name="CheckIfShouldKillVBCSCompiler">
    <CheckIfVBCSCompilerWillOverride src="$(RoslynToolPath)\VBCSCompiler.exe" dest="$(RoslynToolsDestinationFolder)\VBCSCompiler.exe">
      <Output TaskParameter="WillOverride" PropertyName="ShouldKillVBCSCompiler" />
    </CheckIfVBCSCompilerWillOverride>
  </Target>
  <Target Name = "KillVBCSCompilerBeforeCopy" BeforeTargets="CopyRoslynCompilerFilesToOutputDirectory" DependsOnTargets="LocateRoslynToolsDestinationFolder;CheckIfShouldKillVBCSCompiler" >
    <KillProcess ProcessName="VBCSCompiler" ImagePath="$(RoslynToolsDestinationFolder)" Condition="'$(ShouldKillVBCSCompiler)' == 'True'" />
  </Target>
  <Target Name = "KillVBCSCompilerBeforeClean" AfterTargets="BeforeClean" DependsOnTargets="LocateRoslynToolsDestinationFolder">
    <KillProcess ProcessName="VBCSCompiler" ImagePath="$(RoslynToolsDestinationFolder)" />
  </Target>
  <UsingTask TaskName="KillProcess" TaskFactory="CodeTaskFactory" AssemblyFile="$(MSBuildToolsPath)\Microsoft.Build.Tasks.v4.0.dll">
    <ParameterGroup>
      <ProcessName ParameterType="System.String" Required="true" />
      <ImagePath ParameterType="System.String" Required="true" />
    </ParameterGroup>
    <Task>
      <Reference Include="System" />
      <Reference Include="System.Management" />
      <Using Namespace="System" />
      <Using Namespace="System.Linq" />
      <Using Namespace="System.Diagnostics" />
      <Using Namespace="System.Management" />
      <Code Type="Fragment" Language="cs">
        <![CDATA[
                try
                {
                  foreach(var p in Process.GetProcessesByName(ProcessName))
                  {
                      var wmiQuery = "SELECT ProcessId, ExecutablePath FROM Win32_Process WHERE ProcessId = " + p.Id;
                      using(var searcher = new ManagementObjectSearcher(wmiQuery))
                      {
                        using(var results = searcher.Get())
                          {
                            var mo = results.Cast<ManagementObject>().FirstOrDefault();
                            if(mo != null)
                            {
                              var path = (string)mo["ExecutablePath"];
                              var executablePath = path != null ? path : string.Empty;
                              Log.LogMessage("ExecutablePath is {0}", executablePath);

                              if(executablePath.StartsWith(ImagePath, StringComparison.OrdinalIgnoreCase))
                              {
                                p.Kill();
                                p.WaitForExit();
                                Log.LogMessage("{0} is killed", executablePath);
                                break;
                              }
                            }
                          }
                      }
                  }
                }
                catch (Exception ex)
                {
                  Log.LogWarning(ex.Message);
                }
                return true;
                ]]>
      </Code>
    </Task>
  </UsingTask>
  <UsingTask TaskName="CheckIfVBCSCompilerWillOverride" TaskFactory="CodeTaskFactory" AssemblyFile="$(MSBuildToolsPath)\Microsoft.Build.Tasks.v4.0.dll">
    <ParameterGroup>
      <Src ParameterType="System.String" Required="true" />
      <Dest ParameterType="System.String" Required="true" />
      <WillOverride ParameterType="System.Boolean" Output="true" />
    </ParameterGroup>
    <Task>
      <Reference Include="System.IO" />
      <Code Type="Fragment" Language="cs">
        <![CDATA[
                WillOverride = false;
                try {
                  WillOverride = File.Exists(Src) && File.Exists(Dest) && (File.GetLastWriteTime(Src) != File.GetLastWriteTime(Dest));
                } 
                catch { }
                ]]>
      </Code>
    </Task>
  </UsingTask>
</Project>