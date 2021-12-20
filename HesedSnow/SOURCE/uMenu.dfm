object frmMenu: TfrmMenu
  Left = 0
  Top = 0
  Width = 197
  Height = 192
  TabOrder = 0
  TabStop = True
  object btnSLG: TsBitBtn
    AlignWithMargins = True
    Left = 3
    Top = 34
    Width = 191
    Height = 25
    Align = alTop
    Caption = #1057#1051#1043
    TabOrder = 0
    OnClick = btnSLGClick
    ShowFocus = False
  end
  object btnNalogy: TsBitBtn
    AlignWithMargins = True
    Left = 3
    Top = 65
    Width = 191
    Height = 25
    Align = alTop
    Caption = #1055#1086#1076#1072#1090#1082#1080' '#1085#1072' '#1087#1086#1089#1083#1091#1075#1080
    Enabled = False
    TabOrder = 1
    OnClick = btnNalogyClick
    ShowFocus = False
  end
  object btnCreateVidom: TsBitBtn
    AlignWithMargins = True
    Left = 3
    Top = 3
    Width = 191
    Height = 25
    Align = alTop
    Caption = #1057#1090#1074#1086#1088#1080#1090#1080' '#1074#1110#1076#1086#1084#1110#1089#1090#1100
    ModalResult = 1
    TabOrder = 2
    OnClick = btnVidomistClick
    ShowFocus = False
  end
  object btnDavayPodkl: TsBitBtn
    AlignWithMargins = True
    Left = 3
    Top = 96
    Width = 191
    Height = 25
    Align = alTop
    Caption = #1044#1072#1074#1072#1081' '#1087#1110#1076#1082#1083#1102#1095#1080#1084#1089#1103
    TabOrder = 3
    OnClick = btnDavayPodklClick
    ShowFocus = False
  end
  object btnUtils: TsBitBtn
    AlignWithMargins = True
    Left = 3
    Top = 127
    Width = 191
    Height = 25
    Align = alTop
    Caption = #1059#1090#1110#1083#1110#1090#1080
    TabOrder = 4
    OnClick = btnUtilsClick
    ShowFocus = False
  end
  object btnImportBD: TsBitBtn
    AlignWithMargins = True
    Left = 3
    Top = 158
    Width = 191
    Height = 25
    Align = alTop
    Caption = #1030#1084#1087#1086#1088#1090' '#1074' '#1093#1084#1072#1088#1091
    TabOrder = 5
    OnClick = btnImportBDClick
    ShowFocus = False
  end
  object sFrameAdapter1: TsFrameAdapter
    Left = 160
    Top = 40
  end
end
