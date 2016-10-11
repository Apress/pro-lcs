[Version]
Class=IEXPRESS
SEDVersion=3
[Options]
PackagePurpose=InstallApp
ShowInstallProgramWindow=0
HideExtractAnimation=0
UseLongFileName=1
InsideCompressed=0
CAB_FixedSize=0
CAB_ResvCodeSigning=0
RebootMode=N
InstallPrompt=%InstallPrompt%
DisplayLicense=%DisplayLicense%
FinishMessage=%FinishMessage%
TargetName=%TargetName%
FriendlyName=%FriendlyName%
AppLaunched=%AppLaunched%
PostInstallCmd=%PostInstallCmd%
AdminQuietInstCmd=%AdminQuietInstCmd%
UserQuietInstCmd=%UserQuietInstCmd%
SourceFiles=SourceFiles
[Strings]
InstallPrompt=
DisplayLicense=
FinishMessage=
TargetName=C:\work\book\Apress\chapter 6\installCommunicator.EXE
FriendlyName=Communicator Configurator
AppLaunched=msiexec.exe /i Communicator.msi /q 
PostInstallCmd=wscript.exe configure.vbs
AdminQuietInstCmd=
UserQuietInstCmd=
FILE0="Communicator.msi"
FILE1="configure.vbs"
[SourceFiles]
SourceFiles0=c:\temp\
SourceFiles1=C:\work\book\Apress\chapter 6\
[SourceFiles0]
%FILE0%=
[SourceFiles1]
%FILE1%=
