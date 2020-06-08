unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  ComCtrls, ActnList, Menus, StdActns, Buttons,
  fpstypes, fpspreadsheet, fpspreadsheetctrls, fpspreadsheetgrid, fpsActions,
  fpsallformats;

type

  { TForm1 }

  TForm1 = class(TForm)
    ActionList1: TActionList;
    FileExit: TFileExit;
    FileOpen: TFileOpen;
    FileSaveAs: TFileSaveAs;
    ImageList1: TImageList;
    MainMenu1: TMainMenu;
    MnuFile: TMenuItem;
    MnuFileOpen: TMenuItem;
    MnuFileSaveAs: TMenuItem;
    MnuFileSeparator: TMenuItem;
    MnuFileExit: TMenuItem;
    MnuFormat: TMenuItem;
    MnuFormatBold: TMenuItem;
    MnuFormatItalic: TMenuItem;
    MnuFormatUnderline: TMenuItem;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    sHorAlignmentAction1: TsHorAlignmentAction;
    swsSource: TsWorkbookSource;
    swtcWBTab: TsWorkbookTabControl;
    swgSheet: TsWorksheetGrid;
    sceEdit: TsCellEdit;
    sciCell: TsCellIndicator;
    FontBold: TsFontStyleAction;
    FontItalic: TsFontStyleAction;
    FontUnderLine: TsFontStyleAction;
    cbxColour: TsCellCombobox;
    cbxFont: TsCellCombobox;
    cbxSize: TsCellCombobox;
    Splitter3: TSplitter;
    tbToolBar: TToolBar;
    btnOpen: TToolButton;
    btnSave: TToolButton;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    btnBold: TToolButton;
    btnItalic: TToolButton;
    btnUnderline: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    procedure FileOpenAccept(Sender: TObject);
    procedure FileSaveAsAccept(Sender: TObject);
  private

  protected

  public

  end;

var
  Form1: TForm1;


implementation

{$R *.lfm}

uses
  LCLIntf,
  fpsUtils, fpsCSV;


{ TForm1 }

//------------------------------------------------------------------------------
// Called when the user clicks on the Open button after selecting a file
//------------------------------------------------------------------------------
procedure TForm1.FileOpenAccept(Sender: TObject);
begin

  swsSource.AutodetectFormat := false;

  case FileOpen.Dialog.FilterIndex of
    1: swsSource.AutoDetectFormat := true;         // All spreadsheet files
    2: swsSource.AutoDetectFormat := true;         // All Excel files
    3: swsSource.FileFormat := sfOOXML;            // Excel 2007+
    4: swsSource.FileFormat := sfExcel8;           // Excel 97-2003
    5: swsSource.FileFormat := sfExcel5;           // Excel 5.0
    6: swsSource.FileFormat := sfExcel2;           // Excel 2.1
    7: swsSource.FileFormat := sfOpenDocument;     // Open/LibreOffice
    8: swsSource.FileFormat := sfCSV;              // Text files
    9: swsSource.FileFormat := sfHTML;             // HTML files
  end;

  swsSource.FileName := FileOpen.Dialog.FileName;  // This loads the file

end;

//------------------------------------------------------------------------------
// Called when the user clicks on the Save button after selecting a folder and
// output file name
//------------------------------------------------------------------------------
procedure TForm1.FileSaveAsAccept(Sender: TObject);
var
  fmt: TsSpreadsheetFormat;

begin

  Screen.Cursor := crHourglass;

  try

    case FileSaveAs.Dialog.FilterIndex of
      1: fmt := sfOOXML;                // Note: Indexes are 1-based here!
      2: fmt := sfExcel8;
      3: fmt := sfExcel5;
      4: fmt := sfExcel2;
      5: fmt := sfOpenDocument;
      6: fmt := sfCSV;
      7: fmt := sfHTML;
    end;

    swsSource.SaveToSpreadsheetFile(FileSaveAs.Dialog.FileName, fmt);

  finally
    Screen.Cursor := crDefault;
  end;

end;

//------------------------------------------------------------------------------
end.

