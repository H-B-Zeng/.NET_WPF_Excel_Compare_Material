# [.NET][WPF] Excel Compare Material
  
*   [Dev environment](#)
    * [IDE : Visual Studio 2019](#)
    * [DLL : EPPlus](#)
  
Issue:使用調整單的品號，找退料檔裡面出現幾次，並將比對結果匯出成Excel

Point: 如何將WPF打包成單一執行檔(exe)免安裝，步驟如下

[Step1](#AfterResolveReferences):   *.csproj，在最後面加入這一段CODE，將引用的DLL embedded

```
  <Target  Name="AfterResolveReferences">
    <ItemGroup>
      <EmbeddedResource Include="@(ReferenceCopyLocalPaths)" Condition="'%(ReferenceCopyLocalPaths.Extension)' == '.dll'">
        <LogicalName>%(ReferenceCopyLocalPaths.DestinationSubDirectory)%(ReferenceCopyLocalPaths.Filename)%(ReferenceCopyLocalPaths.Extension)</LogicalName>
      </EmbeddedResource>
    </ItemGroup>
  </Target>
```
[Step2](#OnResolveAssembly):修改App.xaml.cs
```
private static Assembly OnResolveAssembly(object sender, ResolveEventArgs args)
        {
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            var executingAssemblyName = executingAssembly.GetName();
            var resName = executingAssemblyName.Name + ".resources";

            AssemblyName assemblyName = new AssemblyName(args.Name); string path = "";
            if (resName == assemblyName.Name)
            {
                path = executingAssemblyName.Name + ".g.resources"; ;
            }
            else
            {
                path = assemblyName.Name + ".dll";
                if (assemblyName.CultureInfo.Equals(CultureInfo.InvariantCulture) == false)
                {
                    path = String.Format(@"{0}\{1}", assemblyName.CultureInfo, path);
                }
            }

            using (Stream stream = executingAssembly.GetManifestResourceStream(path))
            {
                if (stream == null)
                    return null;

                byte[] assemblyRawBytes = new byte[stream.Length];
                stream.Read(assemblyRawBytes, 0, assemblyRawBytes.Length);
                return Assembly.Load(assemblyRawBytes);
            }
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            AppDomain.CurrentDomain.AssemblyResolve += OnResolveAssembly;
        }
```
[Step3](#bulid):rebuild project, export an executable file

