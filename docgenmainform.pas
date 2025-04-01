unit docgenmainform;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, EditBtn, StdCtrls, ExtCtrls, JSONPropStorage;

type

  { TFrmMain }

  TFrmMain = class(TForm)
    BtnGenerate: TButton;
    BtnGenerateList: TButton;
    DrctryOutput: TDirectoryEdit;
    EdtKeyword1: TLabeledEdit;
    EdtNewValue1: TLabeledEdit;
    FlNmEdtValueListFile: TFileNameEdit;
    FlNmEdtTemplateFile: TFileNameEdit;
    EdtKeyword: TLabeledEdit;
    EdtNewValue: TLabeledEdit;
    Label1: TLabel;
    LblOutputDirectory: TLabel;
    PrpStrg: TJSONPropStorage;
    LblTemplateFile: TLabel;
    procedure BtnGenerateClick({%H-}Sender: TObject);
    procedure BtnGenerateListClick({%H-}Sender: TObject);
  private

  public

  end;

var
  FrmMain: TFrmMain;

implementation

uses
  zip_odt, csvreadwrite, odt_2_pdf
  ;



procedure Generate(const aSrcFile: String; aKeyValues: TStringList; const aOutputDir: String);
var
  aODTFile, aNewValue: String;
begin
  if aKeyValues.Count=0 then
    Exit;
  aNewValue:=aKeyValues.ValueFromIndex[0];
  aODTFile:=IncludeTrailingPathDelimiter(aOutputDir)+Trim(aNewValue)+ExtractFileExt(aSrcFile);
  FillODTDoc(aSrcFile, aODTFile, aKeyValues);
  ConvertODT2Pdf(aODTFile, aOutputDir);
end;

procedure Generate(const aSrcFile: String; const aKeywordName, aValueName, aKeywordDate, aValueDate: String;
  const aOutputDir: String);
var
  aKeyValues: TStringList;
begin
  aKeyValues:=TStringList.Create;
  try
    aKeyValues.AddPair(aKeywordName, aValueName);
    aKeyValues.AddPair(aKeywordDate, aValueDate);
    Generate(aSrcFile, aKeyValues, aOutputDir);
  finally
    aKeyValues.Free;
  end;
end;

{$R *.lfm}

{ TFrmMain }

procedure TFrmMain.BtnGenerateClick(Sender: TObject);
begin
  Generate(FlNmEdtTemplateFile.FileName, EdtKeyword.Text, EdtNewValue.Text, EdtKeyword1.Text, EdtNewValue1.Text,
    DrctryOutput.Directory);
end;

procedure TFrmMain.BtnGenerateListClick(Sender: TObject);
var
  aCSV: TCSVParser;
  aFileStream: TFileStream;
  aName: String;
begin
  aCSV:=TCSVParser.Create;
  try
    aFileStream := TFileStream.Create(FlNmEdtValueListFile.FileName, fmOpenRead+fmShareDenyWrite);
    try
      aCSV.SetSource(aFileStream);
      while aCSV.ParseNextCell do
      begin
        if not aCSV.ParseNextCell then
          break;
        if not aCSV.ParseNextCell then
          break;
        aName:=aCSV.CurrentCellText;
        if aCSV.CurrentCol=2 then
          Generate(FlNmEdtTemplateFile.FileName, EdtKeyword.Text, aName, EdtKeyword1.Text, EdtNewValue1.Text,
            DrctryOutput.Directory)
        else
          raise Exception.CreateFmt('Error handle CSV. aName %s, aCol %d', [aName, aCSV.CurrentCol]);
      end;
    finally
      aFileStream.Free;
    end;
  finally
    aCSV.Free;
  end;
end;

end.

