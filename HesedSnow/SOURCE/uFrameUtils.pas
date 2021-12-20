unit uFrameUtils;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls,
  uFrameCustom,
  sFrameAdapter, Vcl.ComCtrls, sMemo, Vcl.Buttons, sBitBtn, Vcl.Grids,
  JvExGrids, JvStringGrid, Data.DB, Vcl.DBGrids, sRichEdit;

type
  TfrmUtils = class(TCustomInfoFrame)
    sBitBtn1: TsBitBtn;
    mInSpisok: TsMemo;
    mOutSpisok: TsRichEdit;
    procedure sBitBtn1Click(Sender: TObject);
  private
    { Private declarations }
  procedure AddColoredLine(ARichEdit: TRichEdit; AText: string; AColor: TColor);

  public
    { Public declarations }
  end;

implementation

{$R *.dfm}

uses uMyProcedure, uDataModul;

procedure TfrmUtils.AddColoredLine(ARichEdit: TRichEdit; AText: string;
  AColor: TColor);
begin
   with ARichEdit do
   begin
     SelStart := Length(Text);
     SelAttributes.Color := AColor;
     SelAttributes.Size := 8;
     SelAttributes.Name := 'MS Sans Serif';
     Lines.Add(AText);
   end;
 end;


procedure TfrmUtils.sBitBtn1Click(Sender: TObject);
var
  i, j: integer;
  str, s, F, N, O, SQLText, txt: String;
  Client: TStringList;

begin
  for i := 0 to mInSpisok.Lines.Count - 1 do
  begin
    s := Trim(mInSpisok.Lines[i]);
    if s <> '' then
    begin
      try
        Client := TStringList.Create;
        // разделяем строку на слова
        Client.text := StringReplace(s, ' ', #13, [rfReplaceAll]);

        if Client.Count = 1 then
        begin
          F := Client[0]; // Familya

          SQLText := 'Like "%' + F + '%"';

        end;

        if Client.Count = 2 then
        begin
          F := Client[0]; // Familya
          N := Client[1]; // Name
          // проверить второе слово или есть точки или две заглавные точки

        //  str := N;
          str := DelSymbolIntoString(N, '.');

          if LastDelimiter('.', str) > 0 then
          begin
            // есть точка
            N := Copy(str, 1, Pos('.', str) - 1);
            O := DelSymbolIntoString(Copy(str, Pos('.', str) + 1), '.');

            SQLText := 'Like "%' + F + ' ' + N + '%"';
          end
          else
          begin
            // если нет точек и длина больше 2 то єто имя
            if str.Length = 2 then
            begin
              O := str[2];
              N := str[1];
            end;
            if str.Length = 1 then N := str[1];

           SQLText := 'Like "%' + F + ' ' + N + '%"';
          end;

        end;

        if Client.Count = 3 then
        begin
          F := Client[0]; // Familya
          N := DelSymbolIntoString(Client[1], '.'); // Name
          O := DelSymbolIntoString(Client[2], '.'); // Otchestvo
          // проверить второе и третье слово или есть точки

          SQLText := 'Like "%' + F + ' ' + N + '%"';

        end;

      finally
        Client.Free;
      end;

    // zapros

      begin
        with DM.qQuery do
        begin
          Active := false;
          Close;
          SQL.Clear;
          SQL.Add('SELECT Uchastniky.[JDC ID], Uchastniky.ФИО FROM Uchastniky WHERE (((Uchastniky.ФИО) '
            + SQLText + '));');

          s := SQL.text;
          Active := true;

          if RecordCount = 1 then
          begin
            mOutSpisok.Lines.Add(FieldbyName('ФИО').asString);
//            mOutSpisok.Lines.Add(FieldbyName('JDC ID').asString + ' ' + FieldbyName('ФИО').asString);
          end;
          if RecordCount > 1 then
          begin
        //  mOutSpisok.Font.Color := clRed;
        //AddColoredLine(ARichEdit: TRichEdit; AText: string;
        // AddColoredLine(RichEdit1, 'Hallo', clRed);
           for j := 0 to RecordCount - 1 do
           begin

           txt := FieldbyName('ФИО').asString;
//           txt := FieldbyName('JDC ID').asString + ' ' + FieldbyName('ФИО').asString;
           AddColoredLine(mOutSpisok, txt, clRed);

            next;

           end;

     //     mOutSpisok.Font.Color := clBlack;

          end;



        end;
      end;

    end;

  end;

end;

end.
