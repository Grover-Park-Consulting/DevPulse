Operation =1
Option =0
Where ="(((IIf(TempVarsLong(\"lngStepID\")=0,0,[tblSession].[StepID])) In (0,TempVarsLon"
    "g(\"lngStepID\"))) AND ((IIf(TempVarsLong(\"lngProjectID\")=0,0,[tblStep].[Proje"
    "ctID])) In (0,TempVarsLong(\"lngProjectID\"))))"
Begin InputTables
    Name ="tblSession"
    Name ="tblStep"
End
Begin OutputColumns
    Expression ="tblSession.SessionID"
    Expression ="tblSession.StepID"
    Expression ="tblSession.SessionDate"
    Expression ="tblSession.SessionStartTime"
    Expression ="tblSession.SessionEndTime"
    Expression ="tblSession.ContextReset"
    Expression ="tblSession.RevisionCycles"
    Expression ="tblSession.Notes"
    Expression ="tblStep.ProjectID"
    Alias ="Expr1"
    Expression ="IIf(TempVarsLong(\"lngProjectID\")=0,0,[tblStep].[ProjectID])"
End
Begin Joins
    LeftTable ="tblSession"
    RightTable ="tblStep"
    Expression ="tblSession.StepID = tblStep.StepID"
    Flag =1
End
Begin OrderBy
    Expression ="tblSession.SessionDate"
    Flag =1
    Expression ="tblSession.SessionStartTime"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "NoFormat" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblSession.RevisionCycles"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSession.StepID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSession.SessionID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSession.SessionDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSession.Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSession.SessionStartTime"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSession.SessionEndTime"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblSession.ContextReset"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblStep.ProjectID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =55
    Top =106
    Right =3766
    Bottom =1804
    Left =-1
    Top =-1
    Right =3683
    Bottom =1058
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =60
        Top =15
        Right =240
        Bottom =195
        Top =0
        Name ="tblSession"
        Name =""
    End
    Begin
        Left =300
        Top =15
        Right =480
        Bottom =195
        Top =0
        Name ="tblStep"
        Name =""
    End
End
