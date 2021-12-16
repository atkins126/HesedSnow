inherited frmNalogy: TfrmNalogy
  Width = 803
  Height = 505
  ExplicitWidth = 803
  ExplicitHeight = 505
  object sPanel2: TsPanel [0]
    Left = 0
    Top = 0
    Width = 803
    Height = 464
    Align = alClient
    TabOrder = 0
    ExplicitWidth = 185
    ExplicitHeight = 41
    object sRollOutPanel1: TsRollOutPanel
      Left = 1
      Top = 1
      Width = 801
      Height = 184
      Align = alTop
      Caption = #1030#1084#1087#1086#1088#1090' '#1074#1110#1076#1086#1084#1086#1089#1090#1110' '#1074' '#1087#1088#1086#1075#1088#1072#1084#1091
      TabOrder = 0
      DesignSize = (
        801
        162)
      object labInfoStatus: TsLabelFX
        Left = 657
        Top = 0
        Width = 128
        Height = 21
        Alignment = taCenter
        Anchors = [akTop, akBottom]
        Caption = '11111111111111111111'
        Color = clBtnFace
        ParentColor = False
        Angle = 0
        Shadow.OffsetKeeper.LeftTop = -3
        Shadow.OffsetKeeper.RightBottom = 5
        Shadow.Mode = smCustom
      end
      object btnImportNalog: TsButton
        Left = 660
        Top = 124
        Width = 125
        Height = 25
        Caption = #1030#1084#1087#1086#1088#1090
        TabOrder = 0
        OnClick = btnImportNalogClick
      end
    end
  end
  object sPanel3: TsPanel [1]
    Left = 0
    Top = 464
    Width = 803
    Height = 41
    Align = alBottom
    TabOrder = 1
    ExplicitTop = 0
    ExplicitWidth = 185
  end
  inherited sFrameAdapter1: TsFrameAdapter
    Left = 16
    Top = 464
  end
  object OpenDialog: TOpenDialog
    Filter = 'xlsx|*.xlsx'
    InitialDir = #1047#1072#1074#1072#1085#1090#1072#1078#1077#1085#1085#1103
    Left = 544
    Top = 445
  end
end
