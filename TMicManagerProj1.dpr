library TMicManagerProj1;

uses
  ComServ,
  TMicManagerProj1_TLB in 'TMicManagerProj1_TLB.pas',
  TMicManagerImpl1 in 'TMicManagerImpl1.pas' {TMicManagerX: TActiveForm} {TMicManagerX: CoClass};

{$E ocx}

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer;

{$R *.TLB}

{$R *.RES}

begin
end.
