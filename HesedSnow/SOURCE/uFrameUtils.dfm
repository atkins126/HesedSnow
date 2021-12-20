inherited frmUtils: TfrmUtils
  Width = 758
  Height = 417
  ExplicitWidth = 758
  ExplicitHeight = 417
  object sBitBtn1: TsBitBtn [0]
    Left = 3
    Top = 389
    Width = 201
    Height = 25
    Caption = #1087#1086#1095#1072#1090#1080' '#1087#1086#1096#1091#1082
    TabOrder = 0
    OnClick = sBitBtn1Click
  end
  object mInSpisok: TsMemo [1]
    Left = 3
    Top = 8
    Width = 214
    Height = 375
    Lines.Strings = (
      #1055#1077#1090#1088#1086#1074#1072' '#1057'.'
      #1044#1077#1085#1077#1075#1072' '#1058'C.'
      #1045#1088#1096#1086#1074#1072' '#1045#1051
      ' '#1044#1077#1085#1077#1075#1072' '#1058'. C.'
      #1044#1077#1085#1077#1075#1072' '#1058'.C.'
      #1043#1077#1088#1074#1080#1094' '#1042#1072#1083#1077#1085#1090#1080#1085#1072
      #1042#1080#1085#1085#1080#1082
      '')
    TabOrder = 1
    Text = 
      #1055#1077#1090#1088#1086#1074#1072' '#1057'.'#13#10#1044#1077#1085#1077#1075#1072' '#1058'C.'#13#10#1045#1088#1096#1086#1074#1072' '#1045#1051#13#10' '#1044#1077#1085#1077#1075#1072' '#1058'. C.'#13#10#1044#1077#1085#1077#1075#1072' '#1058'.C.'#13#10#1043 +
      #1077#1088#1074#1080#1094' '#1042#1072#1083#1077#1085#1090#1080#1085#1072#13#10#1042#1080#1085#1085#1080#1082#13#10#13#10
  end
  object mOutSpisok: TsRichEdit [2]
    Left = 223
    Top = 8
    Width = 522
    Height = 406
    Color = 15921906
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    Zoom = 100
  end
  inherited sFrameAdapter1: TsFrameAdapter
    Left = 744
    Top = 8
  end
end
