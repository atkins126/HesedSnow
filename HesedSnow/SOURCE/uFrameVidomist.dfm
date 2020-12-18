inherited frmVidomost: TfrmVidomost
  Width = 1063
  Height = 705
  ExplicitWidth = 1063
  ExplicitHeight = 705
  object sGradientPanel1: TsGradientPanel [0]
    Left = 0
    Top = 0
    Width = 1063
    Height = 41
    Align = alTop
    Caption = 'sGradientPanel1'
    ShowCaption = False
    TabOrder = 0
    object sButton1: TsButton
      Left = 5
      Top = 8
      Width = 137
      Height = 25
      Caption = #1042#1090#1103#1085#1091#1090#1080' '#1076#1072#1085#1085#1110
      TabOrder = 0
      OnClick = sButton1Click
    end
  end
  object sPanel1: TsPanel [1]
    Left = 0
    Top = 41
    Width = 1063
    Height = 664
    Align = alClient
    Caption = 'sPanel1'
    TabOrder = 1
    object DBGridEhVedomist: TDBGridEh
      Left = 1
      Top = 1
      Width = 1061
      Height = 662
      Align = alClient
      DynProps = <>
      TabOrder = 0
      object RowDetailData: TRowDetailPanelControlEh
      end
    end
  end
  inherited sFrameAdapter1: TsFrameAdapter
    Left = 264
    Top = 8
  end
  object OpenDialog: TOpenDialog
    Left = 360
    Top = 9
  end
end
