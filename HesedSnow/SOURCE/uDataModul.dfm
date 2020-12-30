object DM: TDM
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Height = 250
  Width = 355
  object tVedomost: TADOTable
    Connection = myConnection
    CursorType = ctStatic
    TableName = 'Vidomist'
    Left = 16
    Top = 64
  end
  object myConnection: TADOConnection
    Connected = True
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\GitCopy\HesedSno' +
      'w\BIN\Win32\Debug\Snow.mdb;Persist Security Info=False;'
    LoginPrompt = False
    Mode = cmShareDenyNone
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 16
    Top = 8
  end
  object dsVedomist: TDataSource
    DataSet = tVedomost
    Left = 16
    Top = 112
  end
  object qQuery: TADOQuery
    Connection = myConnection
    Parameters = <>
    Left = 16
    Top = 168
  end
  object ADOTable1: TADOTable
    Left = 304
    Top = 64
  end
  object DataSource1: TDataSource
    Left = 304
    Top = 112
  end
  object ADOQuery1: TADOQuery
    Parameters = <>
    Left = 304
    Top = 168
  end
  object tUslugy: TADOTable
    Connection = myConnection
    CursorType = ctStatic
    TableName = 'Uslugy'
    Left = 80
    Top = 64
  end
  object dsUslugy: TDataSource
    DataSet = qUslugy
    Left = 80
    Top = 168
  end
  object qUslugy: TADOQuery
    Connection = myConnection
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      
        'select Uslugy.[FIO], Uslugy.[JDCID],Uslugy.[RITM],Uslugy.[Number' +
        '],Uslugy.[SABA],Uslugy.[City] from Uslugy')
    Left = 80
    Top = 120
  end
  object dsQslg: TDataSource
    DataSet = qQslg
    Left = 128
    Top = 168
  end
  object qQslg: TADOQuery
    Active = True
    Connection = myConnection
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select * from slg_items')
    Left = 128
    Top = 120
  end
  object tSLG: TADOTable
    Connection = myConnection
    CursorType = ctStatic
    TableName = 'slg_items'
    Left = 128
    Top = 64
    object tSLGName_SLG: TWideStringField
      FieldName = 'Name_SLG'
      Size = 60
    end
    object tSLGActive: TBooleanField
      FieldName = 'Active'
    end
    object tSLGPrice_inch: TBCDField
      FieldName = 'Price_inch'
      Precision = 19
    end
    object tSLGupakovka: TIntegerField
      FieldName = 'upakovka'
    end
  end
end
