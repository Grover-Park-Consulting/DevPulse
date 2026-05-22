Operation =1
Option =0
Where ="(((tblProject.StatusID) In (1,2,3)))"
Begin InputTables
    Name ="tblProject"
    Name ="tlkpStatus"
End
Begin OutputColumns
    Expression ="tblProject.ProjectID"
    Expression ="tblProject.ProjectName"
    Expression ="tlkpStatus.StatusName"
    Expression ="tblProject.StartDate"
End
Begin Joins
    LeftTable ="tblProject"
    RightTable ="tlkpStatus"
    Expression ="tblProject.StatusID = tlkpStatus.StatusID"
    Flag =2
End
Begin OrderBy
    Expression ="tblProject.StartDate"
    Flag =0
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
        dbText "Name" ="tblProject.StartDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblProject.ProjectID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblProject.ProjectName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlkpStatus.StatusName"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =57
    Top =108
    Right =3743
    Bottom =1828
    Left =-1
    Top =-1
    Right =3656
    Bottom =1082
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =64
        Top =16
        Right =256
        Bottom =208
        Top =0
        Name ="tblProject"
        Name =""
    End
    Begin
        Left =670
        Top =159
        Right =862
        Bottom =351
        Top =0
        Name ="tlkpStatus"
        Name =""
    End
End
