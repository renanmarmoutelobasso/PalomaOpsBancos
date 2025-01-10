; Nome do script: PalomaOpsInstaller.iss
[Setup]
AppName=Paloma Ops Bancos
AppVersion=1.0
DefaultDirName=C:\Haarp\Paloma Ops\Bancos
DefaultGroupName=Paloma Ops
UninstallDisplayIcon={app}\ico.ico
Compression=lzma2
SolidCompression=yes
OutputDir=Output
OutputBaseFilename=PalomaOpsInstaller
SetupIconFile=ico.ico

[Files]
; Inclui os arquivos necessários para a instalação
Source: "ico.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "config.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "remessa.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "retorno.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "status.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "tabelas.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "main.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "logo.png"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Ícone na área de trabalho
Name: "{userdesktop}\Paloma Ops Bancos"; Filename: "{app}\main.exe"; WorkingDir: "{app}"; IconFilename: "{app}\ico.ico"
; Ícone no menu iniciar
Name: "{group}\Paloma Ops Bancos"; Filename: "{app}\main.exe"; WorkingDir: "{app}"; IconFilename: "{app}\ico.ico"

[Tasks]
; Criar tarefas adicionais
Name: "CheckFolders"; Description: "Verificar e organizar pastas necessárias"; GroupDescription: "Tarefas adicionais"; Flags: unchecked

[Run]
; Executa o programa após a instalação (opcional)
Filename: "{app}\main.exe"; Description: "Iniciar Paloma Ops Bancos"; Flags: postinstall skipifsilent



[Code]
procedure MoveFilesToRootAndRemoveSubfolders(BaseFolder: String);
var
  SearchRec: TFindRec;
  SourceFile, DestFile: String;
begin
  if FindFirst(BaseFolder + '\*', SearchRec) then
  begin
    try
      repeat
        if (SearchRec.Name <> '.') and (SearchRec.Name <> '..') then
        begin
          SourceFile := BaseFolder + '\' + SearchRec.Name;
          DestFile := ExtractFilePath(BaseFolder) + SearchRec.Name;
          if (SearchRec.Attributes and FILE_ATTRIBUTE_DIRECTORY) = 0 then
          begin
            if RenameFile(SourceFile, DestFile) then
              Log('Arquivo movido: ' + SourceFile + ' para ' + DestFile)
            else
              Log('Erro ao mover o arquivo: ' + SourceFile);
          end;
        end;
      until not FindNext(SearchRec);
    finally
      FindClose(SearchRec);
    end;
    if RemoveDir(BaseFolder) then
      Log('Subpasta removida: ' + BaseFolder)
    else
      Log('Erro ao remover a subpasta: ' + BaseFolder);
  end
  else
    Log('Nenhum arquivo encontrado em: ' + BaseFolder);
end;

procedure InitializeWizard();
var
  i: Integer;
  BaseFolder, RemessaFolder, RetornoFolder: String;
begin

  for i := 1 to 6 do
  begin
    BaseFolder := 'C:\Boletos\Paloma ' + IntToStr(i);
    RemessaFolder := BaseFolder + '\remessa';
    RetornoFolder := BaseFolder + '\retorno';

    if DirExists(RemessaFolder) then
      MoveFilesToRootAndRemoveSubfolders(RemessaFolder)
    else
      Log('Pasta não encontrada: ' + RemessaFolder);

    if DirExists(RetornoFolder) then
      MoveFilesToRootAndRemoveSubfolders(RetornoFolder)
    else
      Log('Pasta não encontrada: ' + RetornoFolder);
  end;
end;
