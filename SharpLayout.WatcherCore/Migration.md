# Migration to .NET 5
Create a new project in Visual Studio.  
In csproj, set:
```
  <PropertyGroup>
    <PreserveCompilationContext>true</PreserveCompilationContext>
  </PropertyGroup>`
```
https://github.com/dotnet/core/issues/2082#issuecomment-442713181  
Build Release configuration.   
Copy all assemblies from `bin\Release\net5.0\refs` folder to embedded resources.  
Update `Watcher.referenceNames` array.  