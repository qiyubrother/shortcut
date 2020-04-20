program ShortCut;

{$APPTYPE CONSOLE}
{$R *.res}

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  StrUtils,
  {这三个单元是必须的}
  ComObj, ActiveX, ShlObj;

procedure CreateLink(ProgramPath, ProgramArg, LinkPath, Descr: String);
var
  AnObj: IUnknown;
  ShellLink: IShellLink;
  AFile: IPersistFile;
  FileName: WideString;
begin
  if UpperCase(ExtractFileExt(LinkPath)) <> '.LNK' then // 检查扩展名是否正确
  begin
    raise Exception.Create('快捷方式的扩展名必须是 lnk!');
    // 若不是则产生异常
  end;
  try
    OleInitialize(nil); // 初始化OLE库，在使用OLE函数前必须调用初始化
    AnObj := CreateComObject(CLSID_ShellLink); // 根据给定的ClassID生成
    // 一个COM对象，此处是快捷方式
    ShellLink := AnObj as IShellLink; // 强制转换为快捷方式接口
    AFile := AnObj as IPersistFile; // 强制转换为文件接口
    // 设置快捷方式属性，此处只设置了几个常用的属性
    ShellLink.SetPath(PChar(ProgramPath)); // 快捷方式的目标文件，一般为可执行文件
    ShellLink.SetArguments(PChar(ProgramArg)); // 目标文件参数
    ShellLink.SetWorkingDirectory(PChar(ExtractFilePath(ProgramPath)));
    // 目标文件的工作目录
    ShellLink.SetDescription(PChar(Descr)); // 对目标文件的描述
    FileName := LinkPath; // 把文件名转换为WideString类型
    AFile.Save(PWChar(FileName), False); // 保存快捷方式
  finally
    OleUninitialize; // 关闭OLE库，此函数必须与OleInitialize成对调用
  end;

end;

procedure OutputError;
begin
  writeln('Use:');
  writeln('  shortcut <app.exe> <app.lnk> [-param:<param>] [-desc:<desc>]');
  writeln('Samples:');
  writeln('  shortcut full_path_app.exe app.lnk');
  writeln('  shortcut full_path_app.exe app.lnk -param:<param>');
  writeln('  shortcut full_path_app.exe app.lnk -desc:<desc>');
  writeln('  shortcut full_path_app.exe app.lnk -param:<param> -desc:<desc>');
  writeln('  shortcut full_path_app.exe app.lnk -desc:<desc> -param:<param>');
end;

var
  appName: string;
  appLnkName: string;
  appParam: string;
  appDesc: string;

begin
  try
    if (paramcount <> 2) and (paramcount <> 3) and (paramcount <> 4) then
    begin
      OutputError;
      // readln;
      exit;
    end;
    if paramcount >= 2 then
    begin
      appName := paramstr(1);
      appLnkName := paramstr(2);
    end;
    if lowercase(rightstr(appLnkName, 3)) <> 'lnk' then
    begin
      writeln('Invalid <app.lnk> name.');
      exit;
    end;

    if (paramcount >= 3) and (MidStr(paramstr(3), 0, 7) = '-param:') then
    begin
      appParam := MidStr(paramstr(3), length('-param:') + 1, length(paramstr(3)) - length('-param:'));
    end
    else if (paramcount >= 3) and (MidStr(paramstr(3), 0, 6) = '-desc:') then
    begin
      appDesc := MidStr(paramstr(3), length('-desc:') + 1, length(paramstr(3)) - length('-desc:'));
    end;

    if (paramcount >= 4) and (MidStr(paramstr(4), 0, 6) = '-desc:') then
    begin
      appDesc := MidStr(paramstr(4), length('-desc:') + 1, length(paramstr(4)) - length('-desc:'));
    end
    else if (paramcount >= 4) and (MidStr(paramstr(4), 0, 7) = '-param:') then
    begin
      appParam := MidStr(paramstr(4), length('-param:') + 1, length(paramstr(4)) - length('-param:'));
    end;

    CreateLink(appName, appParam, appLnkName, appDesc);
    //writeln(appName);
    //writeln(appLnkName);
    //writeln(appParam);
    //writeln(appDesc);
  except
    on E: Exception do
      writeln(E.ClassName, ': ', E.Message);
  end;

end.
