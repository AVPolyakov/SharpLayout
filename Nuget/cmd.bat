copy SharpLayout.nuspec ..\SharpLayout\
NuGet.exe pack ..\SharpLayout\SharpLayout.csproj -Prop Configuration=Release
del ..\SharpLayout\SharpLayout.nuspec

copy SharpLayout.ImageRendering.nuspec ..\SharpLayout.ImageRendering\
NuGet.exe pack ..\SharpLayout.ImageRendering\SharpLayout.ImageRendering.csproj -Prop Configuration=Release
del ..\SharpLayout.ImageRendering\SharpLayout.ImageRendering.nuspec

copy SharpLayout.WatcherCore.nuspec ..\SharpLayout.WatcherCore\
NuGet.exe pack ..\SharpLayout.WatcherCore\SharpLayout.WatcherCore.csproj -Prop Configuration=Release
del ..\SharpLayout.WatcherCore\SharpLayout.WatcherCore.nuspec