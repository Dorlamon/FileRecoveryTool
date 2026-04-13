@echo off
echo ================================
echo Building OfficeEncryptionProbe
echo ================================

dotnet publish OfficeEncryptionProbe.csproj -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true /p:IncludeNativeLibrariesForSelfExtract=true

echo.
echo Build finished.
echo Output:
echo .\bin\Release\net10.0\win-x64\publish\
pause