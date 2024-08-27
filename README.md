# SccmComInterop
A COM interop library for integration with Microsoft Configuration Management (SCCM) 2012.

## Build
### Visual Studio
1. In Solution Explorer, right-click the project and choose Properties.
1. In the side pane, choose Configuration Properties Build and choose Configuration.
1. In the Configuration list at the top, choose Debug or Release.
1. Build the project.

### CLI
1. Navigate to the folder with the .sln file.
1. Use the following dotnet build command, for the required configuration Debug or Release:
```powershell
dotnet build -c <configuration>
```

## Build DLL
After building the project, the generated DLL files can be found under the solution folder, for the required configuration:
```cmd
SccmComInterop\src\SccmComInterop\bin\<configuration>\net6.0
```

## PowerShell
### Define a Microsoft. NET Core class
To define the Microsoft .NET Core class DLL in your PowerShell session, you need to use the `Add-Type` cmdlt:
```powershell
add-type -path <path to Build DLL>\SccmComInterop.dll -PassThru
```

### Getting a list of methods and members:
```powershell
[SccmComInterop.CmInterop].GetMembers() | ft memberType, Name -auto
```

### Using a method

```powershell
 [SccmComInterop.CmInterop]::<methodname>(<params>)
```