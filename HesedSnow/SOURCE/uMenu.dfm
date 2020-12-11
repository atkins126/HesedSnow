object frmMenu: TfrmMenu
  Left = 0
  Top = 0
  Width = 221
  Height = 95
  TabOrder = 0
  object btnVidomist: TsBitBtn
    AlignWithMargins = True
    Left = 3
    Top = 3
    Width = 215
    Height = 25
    Align = alTop
    Caption = #1057#1090#1074#1086#1088#1080#1090#1080' '#1074#1110#1076#1086#1084#1110#1089#1090#1100
    Default = True
    ModalResult = 1
    TabOrder = 0
    OnClick = btnVidomistClick
  end
  object btnSLG: TsBitBtn
    AlignWithMargins = True
    Left = 3
    Top = 34
    Width = 215
    Height = 25
    Align = alTop
    Caption = #1057#1051#1043
    TabOrder = 1
    OnClick = btnSLGClick
  end
  object btnNalogy: TsBitBtn
    AlignWithMargins = True
    Left = 3
    Top = 65
    Width = 215
    Height = 25
    Align = alTop
    Caption = #1055#1086#1076#1072#1090#1082#1080' '#1085#1072' '#1087#1086#1089#1083#1091#1075#1080
    TabOrder = 2
    OnClick = btnNalogyClick
  end
  object sFrameAdapter1: TsFrameAdapter
    Left = 160
    Top = 40
  end
end
