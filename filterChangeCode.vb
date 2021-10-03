Option Compare Database

'So Noticeable stall when you generate. I can probably fix that but it'll have to wait. I know this thing isn't
'nearly as efficient as it could be. There is a lot of duplicated code and more than 1 unnecessary line.
'Anyway it's alive. The last insult access threw at me was VBA interpretting colors backwards. Seriously? Why?

Public db As Database
Public SNum As Integer
Public RNum As Integer
Public RYNum As Integer 'Denoted for overdue items that are only 1 week overdue
Public YNum As Integer
Public SN As Integer
Public RN As Integer
Public RYN As Integer 'Denoted for overdue items that are only 1 week overdue
Public YN As Integer
Public BGRED As Integer
Public BGYREDStart As Integer 'Denoted for overdue items that are only 1 week overdue
Public BGYREDEnd As Integer 'Denoted for overdue items that are only 1 week overdue
Public BGYREDTotal As Integer 'Denoted for overdue items that are only 1 week overdue
Public TemplateDirectory As String
Public NewDateOldDate As Date



Public Sub VarDec()
    
    Set db = CurrentDb
    
End Sub

'I used these variables a lot for a lot things. Sometimes speed sometimes sanity
Public Sub ChangeDec()

    SNum = 1
    RNum = 1
    YNum = 1
    RYNum = 1
    SN = 1
    RN = 1
    RYN = 1
    YN = 1
    BGRED = 0
    BGYREDStart = 0
    BGYREDEnd = 0
    BGYREDTotal = 0
    TemplateDirectory = Application.CurrentProject.Path & "\1.Resources\"
    
End Sub

'Access didn't want to let me join tables in a normal way so I did it this way. 
'Surprisingly this method has never had an issue
Private Sub DateNew_Exit(Cancel As Integer)
    
    VarDec
    
    db.Execute ("UPDATE pg1, pg2, pg3, pg4, pg5, pg6, pg7 " & _
        "SET pg2.DateNew = [pg1].[DateNew], " & _
        "pg3.DateNew = [pg1].[DateNew], " & _
        "pg4.DateNew = [pg1].[DateNew], " & _
        "pg5.DateNew = [pg1].[DateNew], " & _
        "pg6.DateNew = [pg1].[DateNew], " & _
        "pg7.DateNew = [pg1].[DateNew] " & _
        "WHERE pg1.UpdateMarkerDate = FALSE " & _
        "AND pg1.OldDate = pg2.OldDate " & _
        "AND pg1.OldDate = pg3.OldDate " & _
        "AND pg1.OldDate = pg4.OldDate " & _
        "AND pg1.OldDate = pg5.OldDate " & _
        "AND pg1.OldDate = pg6.OldDate " & _
        "AND pg1.OldDate = pg7.OldDate;")

End Sub

'This is The Button. It's on the form and fires off a bunch of subroutines to make the whole thing work.
'I'll talk more about the bits as we continue down
Sub DataEnteredTrigger_Click()

    Dim Confirmation As Integer
 
    Confirmation = MsgBox("Is everything correct?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Data Entry")

    If Confirmation = vbYes Then
        VarDec
        ChangeDec
        OldDataStorage
        InitialCleanUp
        AppendFilterCondition
        TransposeFilterCondition
        DateTransfer
        ConditionTransfer
        ConditionUpdate
        DateChangedUpdate
        DaysUnchanged
        OldRedUpdate
        ScheduledDate
        RequestDate
        BGColor
        CleanUp
        ChangeDateTransfer
        UpdateMarkerDateUpdate
            
    End If

End Sub

'This just moves stuff from last week to the last week column in the table that did a lot of heavy lifting
'Using a table this way was a compromise I had to make with Access to get it to run 
Sub OldDataStorage()
    
    Dim NewDateStorage As DAO.Recordset
    
    Set NewDateStorage = db.OpenRecordset("SELECT pg1.DateNew FROM [pg1] WHERE UpdateMarkerDate = 0 AND AppendMarkerOld = 0")
    NewDateOldDate = NewDateStorage(0).Value
    
    db.Execute ("UPDATE [Filter Condition Data Table] " & _
        "SET [Filter Condition Data Table].OldCondition = [Filter Condition Data Table].CurrentCondition, " & _
        "[Filter Condition Data Table].OldDate = [Filter Condition Data Table].CurrentDate " & _
        "WHERE [Filter Condition Data Table].Exclude = FALSE ")

End Sub

'Cleans up old the table to prevent errors in reports it doesn't cover everything as some of the
'data is needed for decision making later
Sub InitialCleanUp()
    
    db.Execute ("UPDATE [Filter Condition Data Table] " & _
        "SET [Filter Condition Data Table].NumberY = 0, " & _
        "[Filter Condition Data Table].ConditionNumbers = NULL, " & _
        "[Filter Condition Data Table].ConditionWords = NULL, " & _
        "[Filter Condition Data Table].CurrentCondition = NULL " & _
        "WHERE [Filter Condition Data Table].Exclude = FALSE ")
                        
End Sub

'Access refused to display the form with multiple records in a format that was agreeable to the other worker drones
'in my office.  this was an unfortunate compromise to fit 2 week's worth of data onto 1 form. Access also limits you to 255 entries per table
'which is almost never an issue. For this particular use case it was an issue. with 300ish filters
'there were around 3,000 recorded data points. After normalizing it hard I only managed to get it down to around 950
'If I wanted 2 weeks of data it immediately jumped to 1250. This was the beginning of the insanity.
'Anyway this just adds a new record to the table with the data from the previous record shoved into the Old#
'It also puts it in the current section because some pressure gauges don't really change. I'm kind of lazy sometimes.
Sub AppendFilterCondition()

    db.Execute ("INSERT INTO pg1 ( OldDate, Old1, Old2, Old3, Old4, Old5, Old6, Old7, Old8, Old9, Old10, Old11, Old12, Old13, Old14, Old15, Old16, Old17, Old18, Old19, Old20, Old21, Old22, Old23, Old24, Old25, Old26, Old27, Old28, Old29, Old30, Old31, Old32, Old33, Old34, Old35, Old36, Old37, Old38, Old39, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39 ) " & _
        "SELECT DateNew, [1], [2], [3], [4], [5], [6], [7], [8], [9], [10], [11], [12], [13], [14], [15], [16], [17], [18], [19], [20], [21], [22], [23], [24], [25], [26], [27], [28], [29], [30], [31], [32], [33], [34], [35], [36], [37], [38], [39], [1], [2], [3], [4], [5], [6], [7], [8], [9], [10], [11], [12], [13], [14], [15], [16], [17], [18], [19], [20], [21], [22], [23], [24], [25], [26], [27], [28], [29], [30], [31], [32], [33], [34], [35], [36], [37], [38], [39] " & _
        "FROM pg1 " & _
        "WHERE AppendMarkerOld = 0 " & _
        "AND DateNew IS NOT NULL")

    db.Execute ("INSERT INTO pg2 ( OldDate, Old40, Old41, Old42, Old43, Old44, Old45, Old46, Old47, Old48, Old49, Old50, Old51, Old52, Old53, Old54, Old55, Old56, Old57, Old58, Old59, Old60, Old61, Old62, Old63, Old64, Old65, Old66, Old67, Old68, Old69, Old70, Old71, Old72, Old73, Old74, Old75, Old76, Old77, Old78, Old79, Old80, Old81, Old82, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82) " & _
        "SELECT DateNew, [40], [41], [42], [43], [44], [45], [46], [47], [48], [49], [50], [51], [52], [53], [54], [55], [56], [57], [58], [59], [60], [61], [62], [63], [64], [65], [66], [67], [68], [69], [70], [71], [72], [73], [74], [75], [76], [77], [78], [79], [80], [81], [82], [40], [41], [42], [43], [44], [45], [46], [47], [48], [49], [50], [51], [52], [53], [54], [55], [56], [57], [58], [59], [60], [61], [62], [63], [64], [65], [66], [67], [68], [69], [70], [71], [72], [73], [74], [75], [76], [77], [78], [79], [80], [81], [82] " & _
        "FROM pg2 " & _
        "WHERE AppendMarkerOld = 0 " & _
        "AND DateNew IS NOT NULL")
    
    db.Execute ("INSERT INTO pg3 ( OldDate, Old83, Old84, Old85, Old86, Old87, Old88, Old89, Old90, Old91, Old92, Old93, Old94, Old95, Old96, Old97, Old98, Old99, Old100, Old101, Old102, Old103, Old104, Old105, Old106, Old107, Old108, Old109, Old110, Old111, Old112, Old113, Old114, Old115, Old116, Old117, Old118, Old119, Old120, Old121, Old122, Old123, Old124, Old125, Old126, Old127, Old128, Old129, Old130, Old131, Old132, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128, 129, 130, 131, 132 ) " & _
        "SELECT DateNew, [83], [84], [85], [86], [87], [88], [89], [90], [91], [92], [93], [94], [95], [96], [97], [98], [99], [100], [101], [102], [103], [104], [105], [106], [107], [108], [109], [110], [111], [112], [113], [114], [115], [116], [117], [118], [119], [120], [121], [122], [123], [124], [125], [126], [127], [128], [129], [130], [131], [132], [83], [84], [85], [86], [87], [88], [89], [90], [91], [92], [93], [94], [95], [96], [97], [98], [99], [100], [101], [102], [103], [104], [105], [106], [107], [108], [109], [110], [111], [112], [113], [114], [115], [116], [117], [118], [119], [120], [121], [122], [123], [124], [125], [126], [127], [128], [129], [130], [131], [132] " & _
        "FROM pg3 " & _
        "WHERE AppendMarkerOld = 0 " & _
        "AND DateNew IS NOT NULL")

    db.Execute ("INSERT INTO pg4 ( OldDate, Old133, Old134, Old135, Old136, Old137, Old138, Old139, Old140, Old141, Old142, Old143, Old144, Old145, Old146, Old147, Old148, Old149, Old150, Old151, Old152, Old153, Old154, Old155, Old156, Old157, Old158, Old159, Old160, Old161, Old162, Old163, Old164, Old165, Old166, Old167, Old168, Old169, Old170, Old171, Old172, Old173, Old174, Old175, Old176, Old177, Old178, Old179, Old180, 133, 134, 135, 136, 137, 138, 139, 140, 141, 142, 143, 144, 145, 146, 147, 148, 149, 150, 151, 152, 153, 154, 155, 156, 157, 158, 159, 160, 161, 162, 163, 164, 165, 166, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176, 177, 178, 179, 180 ) " & _
        "SELECT DateNew, [133], [134], [135], [136], [137], [138], [139], [140], [141], [142], [143], [144], [145], [146], [147], [148], [149], [150], [151], [152], [153], [154], [155], [156], [157], [158], [159], [160], [161], [162], [163], [164], [165], [166], [167], [168], [169], [170], [171], [172], [173], [174], [175], [176], [177], [178], [179], [180], [133], [134], [135], [136], [137], [138], [139], [140], [141], [142], [143], [144], [145], [146], [147], [148], [149], [150], [151], [152], [153], [154], [155], [156], [157], [158], [159], [160], [161], [162], [163], [164], [165], [166], [167], [168], [169], [170], [171], [172], [173], [174], [175], [176], [177], [178], [179], [180] " & _
        "FROM pg4 " & _
        "WHERE AppendMarkerOld = 0 " & _
        "AND DateNew IS NOT NULL")

    db.Execute ("INSERT INTO pg5 ( OldDate, Old181, Old182, Old183, Old184, Old185, Old186, Old187, Old188, Old189, Old190, Old191, Old192, Old193, Old194, Old195, Old196, Old197, Old198, Old199, Old200, Old201, Old202, Old203, Old204, Old205, Old206, Old207, Old208, Old209, Old210, Old211, Old212, Old213, Old214, Old215, Old216, Old217, Old218, Old219, Old220, Old221, 181, 182, 183, 184, 185, 186, 187, 188, 189, 190, 191, 192, 193, 194, 195, 196, 197, 198, 199, 200, 201, 202, 203, 204, 205, 206, 207, 208, 209, 210, 211, 212, 213, 214, 215, 216, 217, 218, 219, 220, 221 ) " & _
        "SELECT DateNew, [181], [182], [183], [184], [185], [186], [187], [188], [189], [190], [191], [192], [193], [194], [195], [196], [197], [198], [199], [200], [201], [202], [203], [204], [205], [206], [207], [208], [209], [210], [211], [212], [213], [214], [215], [216], [217], [218], [219], [220], [221], [181], [182], [183], [184], [185], [186], [187], [188], [189], [190], [191], [192], [193], [194], [195], [196], [197], [198], [199], [200], [201], [202], [203], [204], [205], [206], [207], [208], [209], [210], [211], [212], [213], [214], [215], [216], [217], [218], [219], [220], [221] " & _
        "FROM pg5 " & _
        "WHERE AppendMarkerOld = 0 " & _
        "AND DateNew IS NOT NULL")

    db.Execute ("INSERT INTO pg6 ( OldDate, Old222, Old223, Old224, Old225, Old226, Old227, Old228, Old229, Old230, Old231, Old232, Old233, Old234, Old235, Old236, Old237, Old238, Old239, Old240, Old241, Old242, Old243, Old244, Old245, Old246, Old247, Old248, Old249, Old250, Old251, Old252, Old253, Old254, Old255, Old256, Old257, Old258, Old259, Old260, Old261, Old262, Old263, Old264, Old265, Old266, Old267, Old268, Old269, Old270, Old271, 222, 223, 224, 225, 226, 227, 228, 229, 230, 231, 232, 233, 234, 235, 236, 237, 238, 239, 240, 241, 242, 243, 244, 245, 246, 247, 248, 249, 250, 251, 252, 253, 254, 255, 256, 257, 258, 259, 260, 261, 262, 263, 264, 265, 266, 267, 268, 269, 270, 271 ) " & _
        "SELECT DateNew, [222], [223], [224], [225], [226], [227], [228], [229], [230], [231], [232], [233], [234], [235], [236], [237], [238], [239], [240], [241], [242], [243], [244], [245], [246], [247], [248], [249], [250], [251], [252], [253], [254], [255], [256], [257], [258], [259], [260], [261], [262], [263], [264], [265], [266], [267], [268], [269], [270], [271], [222], [223], [224], [225], [226], [227], [228], [229], [230], [231], [232], [233], [234], [235], [236], [237], [238], [239], [240], [241], [242], [243], [244], [245], [246], [247], [248], [249], [250], [251], [252], [253], [254], [255], [256], [257], [258], [259], [260], [261], [262], [263], [264], [265], [266], [267], [268], [269], [270], [271] " & _
        "FROM pg6 " & _
        "WHERE AppendMarkerOld = 0 " & _
        "AND DateNew IS NOT NULL")

    db.Execute ("INSERT INTO pg7 ( OldDate, Old272, Old273, Old274, Old275, Old276, Old277, Old278, Old279, Old280, Old281, Old282, Old283, Old284, Old285, Old286, Old287, Old288, Old289, Old290, Old291, Old292, Old293, Old294, Old295, Old296, Old297, Old298, Old299, Old300, 272, 273, 274, 275, 276, 277, 278, 279, 280, 281, 282, 283, 284, 285, 286, 287, 288, 289, 290, 291, 292, 293, 294, 295, 296, 297, 298, 299, 300 ) " & _
        "SELECT DateNew, [272], [273], [274], [275], [276], [277], [278], [279], [280], [281], [282], [283], [284], [285], [286], [287], [288], [289], [290], [291], [292], [293], [294], [295], [296], [297], [298], [299], [300], [272], [273], [274], [275], [276], [277], [278], [279], [280], [281], [282], [283], [284], [285], [286], [287], [288], [289], [290], [291], [292], [293], [294], [295], [296], [297], [298], [299], [300] " & _
        "FROM pg7 " & _
        "WHERE AppendMarkerOld = 0 " & _
        "AND DateNew IS NOT NULL")

    db.Execute ("UPDATE pg1 " & _
        "SET [pg1].AppendMarkerOld = TRUE " & _
        "WHERE [pg1].AppendMarkerOld = FALSE " & _
        "AND [pg1].DateNew IS NOT NULL")

    db.Execute ("UPDATE pg2 " & _
        "SET [pg2].AppendMarkerOld = TRUE " & _
        "WHERE [pg2].AppendMarkerOld = FALSE " & _
        "AND [pg2].DateNew IS NOT NULL")

    db.Execute ("UPDATE pg3 " & _
        "SET [pg3].AppendMarkerOld = TRUE " & _
        "WHERE [pg3].AppendMarkerOld = FALSE " & _
        "AND [pg3].DateNew IS NOT NULL")

    db.Execute ("UPDATE pg4 " & _
        "SET [pg4].AppendMarkerOld = TRUE " & _
        "WHERE [pg4].AppendMarkerOld = FALSE " & _
        "AND [pg4].DateNew IS NOT NULL")

    db.Execute ("UPDATE pg5 " & _
        "SET [pg5].AppendMarkerOld = TRUE " & _
        "WHERE [pg5].AppendMarkerOld = FALSE " & _
        "AND [pg5].DateNew IS NOT NULL")

    db.Execute ("UPDATE pg6 " & _
        "SET [pg6].AppendMarkerOld = TRUE " & _
        "WHERE [pg6].AppendMarkerOld = FALSE " & _
        "AND [pg6].DateNew IS NOT NULL")

    db.Execute ("UPDATE pg7 " & _
        "SET [pg7].AppendMarkerOld = TRUE " & _
        "WHERE [pg7].AppendMarkerOld = FALSE " & _
        "AND [pg7].DateNew IS NOT NULL")
    
End Sub

'Access removed transpose from sql. If you can't get storebought homemade is fine.
'Access also doesn't allow you to make temporary tables. So this thing makes a table
'and transposes a record into a series of records. We'll get back to this in a subroutine or 2 
Sub TransposeFilterCondition()

    db.Execute ("CREATE TABLE FilterGhost (ColumnName VARCHAR, CurrentCondition VARCHAR);")

    Set pg1 = db.OpenRecordset("SELECT pg1.* FROM [pg1] WHERE UpdateMarkerDate = 0")
    Set pg2 = db.OpenRecordset("SELECT pg2.* FROM [pg2] WHERE UpdateMarkerDate = 0")
    Set pg3 = db.OpenRecordset("SELECT pg3.* FROM [pg3] WHERE UpdateMarkerDate = 0")
    Set pg4 = db.OpenRecordset("SELECT pg4.* FROM [pg4] WHERE UpdateMarkerDate = 0")
    Set pg5 = db.OpenRecordset("SELECT pg5.* FROM [pg5] WHERE UpdateMarkerDate = 0")
    Set pg6 = db.OpenRecordset("SELECT pg6.* FROM [pg6] WHERE UpdateMarkerDate = 0")
    Set pg7 = db.OpenRecordset("SELECT pg7.* FROM [pg7] WHERE UpdateMarkerDate = 0")

    For Each A In pg1.Fields
        FN = A.Name
        FC = A.Value
        SQL = "INSERT INTO FilterGhost ([ColumnName],[CurrentCondition]) SELECT '" & FN & "','" & FC & "';"
        db.Execute (SQL)

    Next A

    For Each B In pg2.Fields
        FN = B.Name
        FC = B.Value
        SQL = "INSERT INTO FilterGhost ([ColumnName],[CurrentCondition]) SELECT '" & FN & "','" & FC & "';"
        db.Execute (SQL)

    Next B

    For Each C In pg3.Fields
        FN = C.Name
        FC = C.Value
        SQL = "INSERT INTO FilterGhost ([ColumnName],[CurrentCondition]) SELECT '" & FN & "','" & FC & "';"
        db.Execute (SQL)

    Next C

    For Each D In pg4.Fields
        FN = D.Name
        FC = D.Value
        SQL = "INSERT INTO FilterGhost ([ColumnName],[CurrentCondition]) SELECT '" & FN & "','" & FC & "';"
        db.Execute (SQL)

    Next D

    For Each E In pg5.Fields
        FN = E.Name
        FC = E.Value
        SQL = "INSERT INTO FilterGhost ([ColumnName],[CurrentCondition]) SELECT '" & FN & "','" & FC & "';"
        db.Execute (SQL)

    Next E

    For Each F In pg6.Fields
        FN = F.Name
        FC = F.Value
        SQL = "INSERT INTO FilterGhost ([ColumnName],[CurrentCondition]) SELECT '" & FN & "','" & FC & "';"
        db.Execute (SQL)
        
    Next F

    For Each G In pg7.Fields
        FN = G.Name
        FC = G.Value
        SQL = "INSERT INTO FilterGhost ([ColumnName],[CurrentCondition]) SELECT '" & FN & "','" & FC & "';"
        db.Execute (SQL)

    Next G

End Sub

'Moves the new date from the form onto the heavy lifting work table
Sub DateTransfer()

    db.Execute ("UPDATE pg1, [Filter Condition Data Table] " & _
        "SET [Filter Condition Data Table].CurrentDate = [pg1].[DateNew] " & _
        "WHERE [Filter Condition Data Table].[Marker] = FALSE " & _
        "AND [pg1].[UpdateMarkerDate] = FALSE " & _
        "AND [pg1].[DateNew] IS NOT NULL;")
        
End Sub

'This moves those transposed records from the ghost table to the heavy lifting data table
'Then it deletes the ghost. They literally could have just left the transpose function
'and had a way to make temporary tables. It would have felt less abusive towards baby's first database 
Sub ConditionTransfer()

    'Due to mixed numeric and string types of record data things get a bit crazy.
    db.Execute ("UPDATE [Filter Condition Data Table], FilterGhost " & _
        "SET [Filter Condition Data Table].CurrentCondition = [FilterGhost].[CurrentCondition] " & _
        "WHERE [Filter Condition Data Table].[ColumnName] = [FilterGhost].[ColumnName];")

    db.Execute ("UPDATE [Filter Condition Data Table] " & _
        "SET [Filter Condition Data Table].ConditionWords = [Filter Condition Data Table].CurrentCondition " & _
        "WHERE ISNUMERIC([Filter Condition Data Table].CurrentCondition) = FALSE;")

    db.Execute ("UPDATE [Filter Condition Data Table] " & _
        "SET [Filter Condition Data Table].ConditionNumbers = [Filter Condition Data Table].CurrentCondition " & _
        "WHERE ISNUMERIC([Filter Condition Data Table].CurrentCondition) = TRUE;")

    DoCmd.DeleteObject acTable, "FilterGhost"

End Sub


Sub ConditionUpdate()
    'Tracking for how long items haven't been addressed
    db.Execute ("UPDATE [Filter Condition Data Table] " & _
        "SET NumberR = NumberR + 1 " & _
        "WHERE ConditionWords = RedWords " & _
        "AND ConditionWords IS NOT NULL " & _
        "AND Exclude = FALSE " & _
        "AND Scheduled = FALSE;")
    'This was added due to an error that not even casting the variable as a double could fix.
    'So a quick bit of math and suddenly 0.8 is never greater than 1. That was the only one it happened for.
    'I honestly have no idea, but it's fixed    
    db.Execute ("UPDATE [Filter Condition Data Table] " & _
        "SET [Filter Condition Data Table].NumberR = [Filter Condition Data Table].NumberR + 1 " & _
        "WHERE ([Filter Condition Data Table].ConditionNumbers - [Filter Condition Data Table].RedNumbers) >= 0 " & _
        "AND [Filter Condition Data Table].Exclude = FALSE " & _
        "AND [Filter Condition Data Table].RedNumbers IS NOT NULL " & _
        "AND [Filter Condition Data Table].ConditionNumbers IS NOT NULL;")
    'Clears items from the maintenance list if their number of times needing maintenance is the same this week
    'as it was last week. Last week's number is updated further down.
    db.Execute ("UPDATE [Filter Condition Data Table] " & _
        "SET [Filter Condition Data Table].NumberR = 0 " & _
        "WHERE [Filter Condition Data Table].Exclude = FALSE " & _
        "AND [Filter Condition Data Table].OldRed = [Filter Condition Data Table].NumberR " & _
        "AND [Filter Condition Data Table].OldRed <> 0 ; ")
    'Some item were scheduled for weekly or monthly. Honestly There were better ways to handle those.    
    db.Execute ("UPDATE [Filter Condition Data Table] " & _
        "SET [Filter Condition Data Table].NumberR = 0, " & _
        "[Filter Condition Data Table].OldRed = 0 " & _
        "WHERE [Filter Condition Data Table].Scheduled = TRUE;")
    'This handles items that go on the upcoming maintenance list. It was somewhat tempermental
    db.Execute ("UPDATE [Filter Condition Data Table] " & _
        "SET [Filter Condition Data Table].NumberY = 1 " & _
        "WHERE [Filter Condition Data Table].Exclude = FALSE " & _
        "AND  ([Filter Condition Data Table].ConditionNumbers-[Filter Condition Data Table].YellowNumbers) >= 0 " & _
        "AND [Filter Condition Data Table].NumberR < 1 " & _
        "OR [Filter Condition Data Table].ConditionWords = [Filter Condition Data Table].YellowWords " & _
        "AND [Filter Condition Data Table].CurrentCondition IS NOT NULL " & _
        "AND [Filter Condition Data Table].Scheduled = FALSE;")

End Sub

'This automatically records when items were last changed into the database to track how long they tend to last.
'It cheats a little by taking advantage of all the maintenance being done Sunday and this program being run Monday.
Sub DateChangedUpdate()

    db.Execute ("UPDATE [Filter Condition Data Table] " & _
        "SET [Filter Condition Data Table].DateChanged = DateAdd ('d',-1,[Filter Condition Data Table].CurrentDate) " & _
        "WHERE [Filter Condition Data Table].Exclude = FALSE " & _
        "AND [Filter Condition Data Table].ScheduledDate IS NOT NULL " & _
        "AND [Filter Condition Data Table].OldRed > [Filter Condition Data Table].NumberR " & _
        "AND [Filter Condition Data Table].OldRed <> 0;")

    db.Execute ("UPDATE [Filter Condition Data Table] " & _
        "SET [Filter Condition Data Table].DateChanged = DateAdd ('d',-1,[Filter Condition Data Table].CurrentDate) " & _
        "WHERE [Filter Condition Data Table].Exclude = FALSE " & _
        "AND [Filter Condition Data Table].Scheduled = TRUE ")

End Sub

'This just calculates how long the filter has been in service. Probably could have done that dynamically, but at this point
'I was solo wrestling with access so long that it was just a whatever works moment. 
Sub DaysUnchanged()

    db.Execute ("UPDATE [Filter Condition Data Table] " & _
        "SET [Filter Condition Data Table].DaysUnchanged = DateDiff('d',[Filter Condition Data Table].DateChanged, [Filter Condition Data Table].CurrentDate) " & _
        "WHERE [Filter Condition Data Table].Exclude = FALSE " & _
        "AND [Filter Condition Data Table].DateChanged IS NOT NULL")

End Sub

'Records how many times something has been labelled for maintenance to make calculations easier later. Technically easier earlier.  
Sub OldRedUpdate()

    db.Execute ("UPDATE [Filter Condition Data Table] " & _
        "SET [Filter Condition Data Table].OldRed = [Filter Condition Data Table].NumberR " & _
        "WHERE [Filter Condition Data Table].Exclude = FALSE AND [Filter Condition Data Table].NumberR IS NOT NULL;")
    
End Sub

'Adds the Sunday of the current week as the maintenance day for filters if they have a maintenance value (NumberR) > 1
'Also handles items that are done weekly regardless of status
Sub ScheduledDate()

    db.Execute ("UPDATE [Filter Condition Data Table] " & _
        "SET [Filter Condition Data Table].ScheduledDate = DATEADD ('d', 6, [Filter Condition Data Table].[CurrentDate]) " & _
        "WHERE [Filter Condition Data Table].Exclude = FALSE " & _
        "AND [Filter Condition Data Table].CurrentDate IS NOT NULL " & _
        "AND  [Filter Condition Data Table].NumberR >= 1 " & _
        "AND [Filter Condition Data Table].Scheduled = FALSE;")

    db.Execute ("UPDATE [Filter Condition Data Table] " & _
        "SET [Filter Condition Data Table].ScheduledDate = [Filter Condition Data Table].CurrentCondition " & _
        "WHERE [Filter Condition Data Table].Exclude = FALSE " & _
        "AND [Filter Condition Data Table].CurrentDate IS NOT NULL " & _
        "AND [Filter Condition Data Table].Scheduled = TRUE;")

End Sub

'Records the current date as the day a request for maintenance was sent out
Sub RequestDate()

    db.Execute ("UPDATE [Filter Condition Data Table] " & _
        "SET [Filter Condition Data Table].RequestDate = [Filter Condition Data Table].CurrentDate " & _
        "WHERE [Filter Condition Data Table].Exclude = FALSE " & _
        "AND [Filter Condition Data Table].Scheduled = FALSE " & _
        "AND  [Filter Condition Data Table].NumberR = 1;")

    db.Execute ("UPDATE [Filter Condition Data Table] " & _
        "SET [Filter Condition Data Table].RequestDate = [Filter Condition Data Table].CurrentCondition " & _
        "WHERE [Filter Condition Data Table].Exclude = FALSE " & _
        "AND [Filter Condition Data Table].CurrentCondition IS NOT NULL " & _
        "AND [Filter Condition Data Table].Scheduled = TRUE;")

End Sub

'The really funny part is how long it took to make this and it isn't even used in the final
'I kept it as a monument to my frustation. And to possibly repurpose if requirements for coloring were changed. 
Sub BGColor()

    db.Execute ("UPDATE [Filter Condition Data Table] " & _
        "SET [Filter Condition Data Table].BGColor = '16777215' " & _
        "WHERE [Filter Condition Data Table].NumberR <=1 " & _
        "OR [Filter Condition Data Table].Scheduled = True;")
        
    db.Execute ("UPDATE [Filter Condition Data Table] " & _
        "SET [Filter Condition Data Table].BGColor = '65535' " & _
        "WHERE [Filter Condition Data Table].NumberR = 2 " & _
        "AND [Filter Condition Data Table].Scheduled = FALSE;")
'Hopefully the Above will set the BGColor note to yellow on these items. VBA reversing long for RGB messes with me

    db.Execute ("UPDATE [Filter Condition Data Table] " & _
        "SET [Filter Condition Data Table].BGColor = '255' " & _
        "WHERE [Filter Condition Data Table].NumberR > 2 " & _
        "AND [Filter Condition Data Table].Scheduled = FALSE;")

End Sub

'Stage 2 of cleanup to prevent issues later. It couldn't be done earlier as we needed that number
Sub CleanUp()
    
        db.Execute ("UPDATE [Filter Condition Data Table] " & _
            "SET [Filter Condition Data Table].RequestDate = NULL, " & _
            "[Filter Condition Data Table].ScheduledDate = NULL " & _
            "WHERE [Filter Condition Data Table].NumberR = 0 " & _
            "AND [Filter Condition Data Table].Exclude = FALSE " & _
            "AND [Filter Condition Data Table].Scheduled = FALSE;")
            
End Sub

'This basically detransposes data from the heavy lifting table back to the form tables for long term storage
Sub ChangeDateTransfer()

    For A = 1 To 39
        db.Execute ("UPDATE pg1, [Filter Condition Data Table] " & _
            "SET pg1.Days" & A & " = [Filter Condition Data Table].DaysUnchanged, " & _
            "pg1.Changed" & A & " = [Filter Condition Data Table].DateChanged " & _
            "WHERE [Filter Condition Data Table].ColumnName = '" & A & "' " & _
            "AND AppendMarkerOld = TRUE " & _
            "AND UpdateMarkerDate = FALSE;")
    
    Next A

    For B = 40 To 82
        db.Execute ("UPDATE pg2, [Filter Condition Data Table] " & _
            "SET pg2.Days" & B & " = [Filter Condition Data Table].DaysUnchanged, " & _
            "pg2.Changed" & B & " = [Filter Condition Data Table].DateChanged " & _
            "WHERE [Filter Condition Data Table].ColumnName = '" & B & "' " & _
            "AND AppendMarkerOld = TRUE " & _
            "AND UpdateMarkerDate = FALSE;")

    Next B

    For C = 83 To 132
        db.Execute ("UPDATE pg3, [Filter Condition Data Table] " & _
            "SET pg3.Days" & C & " = [Filter Condition Data Table].DaysUnchanged, " & _
            "pg3.Changed" & C & " = [Filter Condition Data Table].DateChanged " & _
            "WHERE [Filter Condition Data Table].ColumnName = '" & C & "' " & _
            "AND AppendMarkerOld = TRUE " & _
            "AND UpdateMarkerDate = FALSE;")

    Next C

    For D = 133 To 180
        db.Execute ("UPDATE pg4, [Filter Condition Data Table] " & _
            "SET pg4.Days" & D & " = [Filter Condition Data Table].DaysUnchanged, " & _
            "pg4.Changed" & D & " = [Filter Condition Data Table].DateChanged " & _
            "WHERE [Filter Condition Data Table].ColumnName = '" & D & "' " & _
            "AND AppendMarkerOld = TRUE " & _
            "AND UpdateMarkerDate = FALSE;")

    Next D

    For E = 181 To 221
        db.Execute ("UPDATE pg5, [Filter Condition Data Table] " & _
            "SET pg5.Days" & E & " = [Filter Condition Data Table].DaysUnchanged, " & _
            "pg5.Changed" & E & " = [Filter Condition Data Table].DateChanged " & _
            "WHERE [Filter Condition Data Table].ColumnName = '" & E & "' " & _
            "AND AppendMarkerOld = TRUE " & _
            "AND UpdateMarkerDate = FALSE;")

    Next E

    For F = 222 To 271
        db.Execute ("UPDATE pg6, [Filter Condition Data Table] " & _
            "SET pg6.Days" & F & " = [Filter Condition Data Table].DaysUnchanged, " & _
            "pg6.Changed" & F & " = [Filter Condition Data Table].DateChanged " & _
            "WHERE [Filter Condition Data Table].ColumnName = '" & F & "' " & _
            "AND AppendMarkerOld = TRUE " & _
            "AND UpdateMarkerDate = FALSE;")
    
    Next F

    For G = 272 To 300
        db.Execute ("UPDATE pg7, [Filter Condition Data Table] " & _
            "SET pg7.Days" & G & " = [Filter Condition Data Table].DaysUnchanged, " & _
            "pg7.Changed" & G & " = [Filter Condition Data Table].DateChanged " & _
            "WHERE [Filter Condition Data Table].ColumnName = '" & G & "' " & _
            "AND AppendMarkerOld = TRUE AND " & _
            "UpdateMarkerDate = FALSE;")

    Next G

End Sub

'This is the other big button. It generates the maintenance request.
'It has speed issues. I had an idea to make it faster, but never got to try it
Private Sub FilterChangeReq_Click()

    VarDec
  'This was a later addition due to the difficulty of fixing the database if this button is pressed prematurely    
    ProceedCheck = MsgBox("This will cause the database to freeze while it processes. Do not close Access until generation is complete. Generate change out request?", vbQuestion + vbYesNo + vbDefaultButton2, "Proceed")

    If ProceedCheck = vbYes Then
        FilterChangeOutRequestUpdate
    End If
    
End Sub


Sub FilterChangeOutRequestUpdate()

    ChangeDec
    ScheduledChangeOut
    RedChangeOut
    YellowChangeOut
    ChangeOutExport

End Sub

'This thing basically constructs an excel file in a table.  
Sub ScheduledChangeOut()
    'CLEANSE THE TABLE!!!!!!
    db.Execute ("DELETE * FROM [Filter Change Out Request];")
    'Adding a header line for the regularly scheduled items
    db.Execute ("INSERT INTO [Filter Change Out Request] ([Item#], [BGColor]) " & _
        "SELECT 'Weekly Items', '16777215';")
    'Column names for the form
    db.Execute ("INSERT INTO [Filter Change Out Request] ( [Item#], [Priority], [RequestDate], [Location/ID], [BayLocation(Map)], [FilterType], [StoresNumber&Size], [StoresNumber], [QuantityNeeded], [DuringProduction], [ScheduledDate], [ReasonForChange], [TeamSignOff], [BGColor]) " & _
        "SELECT 'Item #', 'Priority', 'Request Date', 'Location/ID', 'Bay Location (Map)', 'Filter Type', 'Stores Number & Size', 'Stores Number', 'Quantity Needed', 'During Production Yes or No', 'Scheduled Date For Change-Out', 'Reason for Change Out-Additional Information', 'TEAM-Sign Off When Completed (Notes)', '13434879';")
    'Finds scheduled items
    Set Scheduled = db.OpenRecordset("SELECT [ColumnTempName], [Priority], [RequestDate], [Location/ID], [BayLocation(Map)], [FilterType], [StoresNumber&Size], [StoresNumber], [QuantityNeeded], [DuringProduction], [ScheduledDate], [ReasonForChange], [BGColor] " & _
        "FROM [Filter Condition Data Table] " & _
        "WHERE Scheduled = TRUE " & _
        "AND Exclude = FALSE " & _
        "AND Combined = FALSE " & _
        "ORDER BY PrimaryKey")
    'Numbers the Scheduleditems
    Do While Not Scheduled.EOF
        Scheduled.Edit
        Scheduled("ColumnTempName") = SNum
        SNum = SNum + 1
        Scheduled.Update
        Scheduled.MoveNext
    Loop
    
    If Scheduled.RecordCount <> 0 Then
        Scheduled.MoveFirst
    End If
    'Moves the scheduled items in order of number.     
    Do While Not Scheduled.EOF
        db.Execute ("INSERT INTO [Filter Change Out Request] ( [Item#], [Priority], [RequestDate], [Location/ID], [BayLocation(Map)], [FilterType], [StoresNumber&Size], [StoresNumber], [QuantityNeeded], [DuringProduction], [ScheduledDate], [ReasonForChange], [TeamSignOff], [BGColor]) " & _
            "SELECT [ColumnTempName], [Priority], [RequestDate], [Location/ID], [BayLocation(Map)], [FilterType], [StoresNumber&Size], [StoresNumber], [QuantityNeeded], [DuringProduction], [ScheduledDate], [ReasonForChange], ' ', [BGColor] " & _
            "FROM [Filter Condition Data Table] " & _
            "WHERE Scheduled = TRUE " & _
            "AND Exclude = FALSE " & _
            "AND Combined = FALSE " & _
            "AND [ColumnTempName] = '" & SN & "' ;")
        SN = SN + 1
        Scheduled.MoveNext
    Loop

End Sub

'This thing is a thing. Basically works the same as the one above except it does it 3 times
'Once for items on the maintenance list for 2 or more weeks and they come first
'Once for items that have been on it 1 week and then all the new stuff that was added this week
Sub RedChangeOut()

    db.Execute ("INSERT INTO [Filter Change Out Request] ([Item#], [BGColor]) " & _
        "SELECT 'Shutdown Items', '16777215';")

    db.Execute ("INSERT INTO [Filter Change Out Request] ( [Item#], [Priority], [RequestDate], [Location/ID], [BayLocation(Map)], [FilterType], [StoresNumber&Size], [StoresNumber], [QuantityNeeded], [DuringProduction], [ScheduledDate], [ReasonForChange], [TeamSignOff], [BGColor]) " & _
        "SELECT 'Item #', 'Priority', 'Request Date', 'Location/ID', 'Bay Location (Map)', 'Filter Type', 'Stores Number & Size', 'Stores Number', 'Quantity Needed', 'During Production Yes or No', 'Scheduled Date For Change-Out', 'Reason for Change Out-Additional Information', 'TEAM-Sign Off When Completed (Notes)', '13434879';")
    
    Set RedOut = db.OpenRecordset("SELECT [ColumnTempName], [Priority], [RequestDate], [Location/ID], [BayLocation(Map)], [FilterType], [StoresNumber&Size], [StoresNumber], [QuantityNeeded], [DuringProduction], [ScheduledDate], [ReasonForChange], [BGColor] " & _
        "FROM [Filter Condition Data Table] " & _
        "WHERE Scheduled = FALSE " & _
        "AND Exclude = FALSE " & _
        "AND NumberR > 2 " & _
        "ORDER BY PrimaryKey")

    Do While Not RedOut.EOF
        RedOut.Edit
        RedOut("ColumnTempName") = RNum
        RNum = RNum + 1
        RedOut.Update
        RedOut.MoveNext
    Loop
    
    BGRED = RNum - 1 'Sneaky global variable to note where red overdue items end
'I did this to try to make it run faster on the copy to excel. It was unsuccessful, but it works as well as my other method    
    BGYREDStart = RNum 'Sneaky global variable to note where yellow items start
    
    If RedOut.RecordCount <> 0 Then
        RedOut.MoveFirst
    End If
         
    Do While Not RedOut.EOF
        db.Execute ("INSERT INTO [Filter Change Out Request] ( [Item#], [Priority], [RequestDate], [Location/ID], [BayLocation(Map)], [FilterType], [StoresNumber&Size], [StoresNumber], [QuantityNeeded], [DuringProduction], [ScheduledDate], [ReasonForChange], [TeamSignOff], [BGColor]) " & _
            "SELECT [ColumnTempName], [Priority], [RequestDate], [Location/ID], [BayLocation(Map)], [FilterType], [StoresNumber&Size], [StoresNumber], [QuantityNeeded], [DuringProduction], [ScheduledDate], [ReasonForChange], ' ', [BGColor] " & _
            "FROM [Filter Condition Data Table] " & _
            "WHERE Scheduled = FALSE " & _
            "AND Exclude = FALSE " & _
            "AND NumberR > 2 " & _
            "AND [ColumnTempName] = '" & RN & "' ;")
        RN = RN + 1
        RedOut.MoveNext
    Loop

'End Of 2+ weeks overdue

    Set RedOut = db.OpenRecordset("SELECT [ColumnTempName], [Priority], [RequestDate], [Location/ID], [BayLocation(Map)], [FilterType], [StoresNumber&Size], [StoresNumber], [QuantityNeeded], [DuringProduction], [ScheduledDate], [ReasonForChange], [BGColor] " & _
        "FROM [Filter Condition Data Table] " & _
        "WHERE Scheduled = FALSE " & _
        "AND Exclude = FALSE " & _
        "AND NumberR = 2 " & _
        "ORDER BY PrimaryKey")
    
    Do While Not RedOut.EOF
        RedOut.Edit
        RedOut("ColumnTempName") = RNum
        RNum = RNum + 1
        RedOut.Update
        RedOut.MoveNext
    Loop
    
    BGYREDEnd = RNum - 1  'Sneaky global variable to note where yellow items End
    
    BGYREDTotal = RNum - BGYREDStart  'Sneaky global variable to note total Yellow overdue items
    
    If RedOut.RecordCount <> 0 Then
        RedOut.MoveFirst
    End If
    
    Do While Not RedOut.EOF
        db.Execute ("INSERT INTO [Filter Change Out Request] ( [Item#], [Priority], [RequestDate], [Location/ID], [BayLocation(Map)], [FilterType], [StoresNumber&Size], [StoresNumber], [QuantityNeeded], [DuringProduction], [ScheduledDate], [ReasonForChange], [TeamSignOff], [BGColor]) " & _
            "SELECT [ColumnTempName], [Priority], [RequestDate], [Location/ID], [BayLocation(Map)], [FilterType], [StoresNumber&Size], [StoresNumber], [QuantityNeeded], [DuringProduction], [ScheduledDate], [ReasonForChange], ' ', [BGColor] " & _
            "FROM [Filter Condition Data Table] " & _
            "WHERE Scheduled = FALSE " & _
            "AND Exclude = FALSE " & _
            "AND NumberR = 2 " & _
            "AND [ColumnTempName] = '" & RN & "' ;")
        RN = RN + 1
        RedOut.MoveNext
    Loop
        
    If RedOut.RecordCount <> 0 Then
        RedOut.MoveFirst
    End If
    
'End Of 1 weeks overdue

    Set RedOut = db.OpenRecordset("SELECT [ColumnTempName], [Priority], [RequestDate], [Location/ID], [BayLocation(Map)], [FilterType], [StoresNumber&Size], [StoresNumber], [QuantityNeeded], [DuringProduction], [ScheduledDate], [ReasonForChange], [BGColor] " & _
        "FROM [Filter Condition Data Table] " & _
        "WHERE Scheduled = FALSE " & _
        "AND Exclude = FALSE " & _
        "AND NumberR = 1 " & _
        "ORDER BY PrimaryKey")
    
    Do While Not RedOut.EOF
        RedOut.Edit
        RedOut("ColumnTempName") = RNum
        RNum = RNum + 1
        RedOut.Update
        RedOut.MoveNext
    Loop
    
    If RedOut.RecordCount <> 0 Then
        RedOut.MoveFirst
    End If
    
    Do While Not RedOut.EOF
        db.Execute ("INSERT INTO [Filter Change Out Request] ( [Item#], [Priority], [RequestDate], [Location/ID], [BayLocation(Map)], [FilterType], [StoresNumber&Size], [StoresNumber], [QuantityNeeded], [DuringProduction], [ScheduledDate], [ReasonForChange], [TeamSignOff], [BGColor]) " & _
            "SELECT [ColumnTempName], [Priority], [RequestDate], [Location/ID], [BayLocation(Map)], [FilterType], [StoresNumber&Size], [StoresNumber], [QuantityNeeded], [DuringProduction], [ScheduledDate], [ReasonForChange], ' ', [BGColor] " & _
            "FROM [Filter Condition Data Table] " & _
            "WHERE Scheduled = FALSE " & _
            "AND Exclude = FALSE " & _
            "AND NumberR = 1 " & _
            "AND [ColumnTempName] = '" & RN & "' ;")
        RN = RN + 1
        RedOut.MoveNext
    Loop
    
'End of new replacement items
    
End Sub

'The same as the others but deals in items that are upcoming
Sub YellowChangeOut()
 
    db.Execute ("INSERT INTO [Filter Change Out Request] ([Item#], [BGColor]) " & _
        "SELECT 'Upcoming Items', '16777215';")
    
    db.Execute ("INSERT INTO [Filter Change Out Request] ( [Item#], [Priority], [RequestDate], [Location/ID], [BayLocation(Map)], [FilterType], [StoresNumber&Size], [StoresNumber], [QuantityNeeded], [DuringProduction], [ScheduledDate], [ReasonForChange], [TeamSignOff], [BGColor]) " & _
        "SELECT 'Item #', 'Priority', 'Request Date', 'Location/ID', 'Bay Location (Map)', 'Filter Type', 'Stores Number & Size', 'Stores Number', 'Quantity Needed', 'During Production Yes or No', 'Scheduled Date For Change-Out', 'Reason for Change Out-Additional Information', 'TEAM-Sign Off When Completed (Notes)', '13434879';")

    Set YellowOut = db.OpenRecordset("SELECT [ColumnTempName], [Priority], [RequestDate], [Location/ID], [BayLocation(Map)], [FilterType], [StoresNumber&Size], [StoresNumber], [QuantityNeeded], [DuringProduction], [ScheduledDate], [ReasonForChange], [BGColor] " & _
        "FROM [Filter Condition Data Table] " & _
        "WHERE Scheduled = FALSE " & _
        "AND Exclude = FALSE " & _
        "AND NumberY = '1' " & _
        "ORDER BY PrimaryKey")
    
    Do While Not YellowOut.EOF
        YellowOut.Edit
        YellowOut("ColumnTempName") = YNum
        YNum = YNum + 1
        YellowOut.Update
        YellowOut.MoveNext
    Loop
    
    If YellowOut.RecordCount <> 0 Then
        YellowOut.MoveFirst
    End If
     
    Do While Not YellowOut.EOF
        db.Execute ("INSERT INTO [Filter Change Out Request] ( [Item#], [Priority], [RequestDate], [Location/ID], [BayLocation(Map)], [FilterType], [StoresNumber&Size], [StoresNumber], [QuantityNeeded], [DuringProduction], [ScheduledDate], [ReasonForChange], [TeamSignOff], [BGColor]) " & _
            "SELECT [ColumnTempName], [Priority], [RequestDate], [Location/ID], [BayLocation(Map)], [FilterType], [StoresNumber&Size], [StoresNumber], [QuantityNeeded], [DuringProduction], [ScheduledDate], [ReasonForChange], ' ', [BGColor] " & _
            "FROM [Filter Condition Data Table] " & _
            "WHERE Scheduled = FALSE " & _
            "AND Exclude = FALSE " & _
            "AND NumberY = '1' " & _
            "AND [ColumnTempName] = '" & YN & "' ;")
        YN = YN + 1
        YellowOut.MoveNext
    Loop

End Sub

'This thing copies all the data from the table above into and excel document for printing
Sub ChangeOutExport()
    
    Dim ChangeRequest As String
    Dim FilterChangeTable As DAO.Recordset
    Dim KCAPTFCOR As Excel.Workbook
    Dim COR As Excel.Worksheet
    Dim CRow As Integer
    Dim CCol As Integer
    Dim ColTitleRowF As String
    Dim ColTitleRowS As String
    Dim ColTitleRowL As String
    Dim SecTitleRowF As String
    Dim SecTitleRowS As String
    Dim SecTitleRowL As String
    Dim SaveDatePDF As DAO.Recordset
    Dim CleanDate As String
    
    ChangeRequest = TemplateDirectory & "KCAPTFCOR.xlsx"
    Set FilterChangeTable = db.OpenRecordset("SELECT * FROM [Filter Change Out Request]")
    Set KCAPTFCOR = Excel.Workbooks.Open(ChangeRequest)
    Set COR = KCAPTFCOR.Worksheets("Change Out Request")
    CRow = 1
    CCol = 1
    'SN, RN, and YN are global variables that were used above for the loop function to number items on the list. Recycling
    ColTitleRowF = "A2:M2"
    ColTitleRowS = "A" & (3 + SN) & ":M" & (3 + SN)
    ColTitleRowL = "A" & (4 + SN + RN) & ":M" & (4 + SN + RN)
    SecTitleRowF = "A1:k1"
    SecTitleRowS = "A" & (2 + SN) & ":M" & (2 + SN)
    SecTitleRowL = "A" & (3 + SN + RN) & ":M" & (3 + SN + RN)
        
    COR.Range("A1:M125").Value = ""
    COR.Range("A1:M125").Font.FontStyle = "Regular"
    COR.Range("A1:M125").Font.Underline = xlUnderlineStyleNone
    COR.Range("A1:M125").Font.Size = 14
    COR.Range("A1:M125").Font.Color = 0
    COR.Range("A1:M125").Borders.LineStyle = xlNone
    COR.Range("A1:M125").Interior.ColorIndex = 0
    COR.Range("A1:M125").HorizontalAlignment = xlHAlignCenter
    COR.Range("A1:M125").UnMerge
    
    FilterChangeTable.MoveFirst
      
    Do While CRow < (SN + RN + YN + 4)

        Do While CCol < 14
            COR.Cells(CRow, CCol).Value = FilterChangeTable(CCol).Value
            CCol = CCol + 1
        Loop
          
        If FilterChangeTable.EOF = False Then
            FilterChangeTable.MoveNext
        End If

        CCol = 1
        CRow = CRow + 1
    
    Loop
'Color codes items on the maintenance form by how long they've been on the list    
    If BGRED > 0 Then
        COR.Range("A" & (4 + SN) & ":M" & (3 + SN + BGRED)).Interior.Color = 255
    End If
      
    If BGYREDTotal > 0 Then
        COR.Range("A" & (4 + SN + BGRED) & ":M" & (3 + SN + BGYREDEnd)).Interior.Color = 65535
    End If
    
    COR.Range("k3:l" & (1 + SN)).Font.Color = 255
    COR.Range("A2:M" & (SN + RN + YN + 3)).Borders.LineStyle = xlDash
    
    COR.Range("L1").Value = "Prepared Date"
    COR.Range("M1").Value = Date 'Quick addition, by request, to have the week a file was prepared in visible on the printout
'This is the almighty formatting chunk. It formats chunks        
    COR.Range(SecTitleRowF).Font.FontStyle = "Bold"
    COR.Range(SecTitleRowS).Font.FontStyle = "Bold"
    COR.Range(SecTitleRowL).Font.FontStyle = "Bold"
    COR.Range(SecTitleRowF).Borders.LineStyle = xlNone
    COR.Range(SecTitleRowS).Borders.LineStyle = xlNone
    COR.Range(SecTitleRowL).Borders.LineStyle = xlNone
    COR.Range(SecTitleRowF).Font.Size = 16
    COR.Range(SecTitleRowS).Font.Size = 16
    COR.Range(SecTitleRowL).Font.Size = 16
    COR.Range(SecTitleRowF).Font.Color = 15773696
    COR.Range(SecTitleRowS).Font.Color = 15773696
    COR.Range(SecTitleRowL).Font.Color = 15773696
    COR.Range(SecTitleRowF).Merge
    COR.Range(SecTitleRowS).Merge
    COR.Range(SecTitleRowL).Merge
    COR.Range(SecTitleRowF).HorizontalAlignment = xlLeft
    COR.Range(SecTitleRowS).HorizontalAlignment = xlLeft
    COR.Range(SecTitleRowL).HorizontalAlignment = xlLeft
    
    COR.Range(ColTitleRowF).Font.FontStyle = "Bold Italic"
    COR.Range(ColTitleRowS).Font.FontStyle = "Bold Italic"
    COR.Range(ColTitleRowL).Font.FontStyle = "Bold Italic"
    COR.Range(ColTitleRowF).Font.Underline = xlUnderlineStyleSingle
    COR.Range(ColTitleRowS).Font.Underline = xlUnderlineStyleSingle
    COR.Range(ColTitleRowL).Font.Underline = xlUnderlineStyleSingle
    COR.Range(ColTitleRowF).Font.Size = 16
    COR.Range(ColTitleRowS).Font.Size = 16
    COR.Range(ColTitleRowL).Font.Size = 16
    COR.Range(ColTitleRowF).Font.Color = 6299648
    COR.Range(ColTitleRowS).Font.Color = 6299648
    COR.Range(ColTitleRowL).Font.Color = 6299648
    COR.Range(ColTitleRowF).Borders.LineStyle = xlContinuous
    COR.Range(ColTitleRowS).Borders.LineStyle = xlContinuous
    COR.Range(ColTitleRowL).Borders.LineStyle = xlContinuous
    COR.Range(ColTitleRowF).Interior.Color = 13434879
    COR.Range(ColTitleRowS).Interior.Color = 13434879
    COR.Range(ColTitleRowL).Interior.Color = 13434879
    'It saves a copy as a pdf and an excel file in case you need to make manual changes
    Set SaveDatePDF = db.OpenRecordset("SELECT [CurrentDate] FROM [Filter Condition Data Table]")
    
    SaveDatePDF.MoveFirst
    'Also names the file after the current date    
    CleanDate = Replace(CStr(SaveDatePDF(0).Value), "/", "")
    
    'I shouldn't have needed to have done it this way, but access was being evil
    Path = Application.CurrentProject.Path
    ChDir Path
    ChDir ".."
    ChDir ".."
    
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, _
    FileName:=CurDir & "\Filter Program\1. Filter Condition\Auto Generated\Filter Change Request" & CleanDate & ".pdf"
        
    ActiveWorkbook.SaveAs FileName:=CurDir & "\Filter Program\1. Filter Condition\Auto Generated\Filter Change Request" & CleanDate & ".xlsx"
        
    'ActiveWorkbook.PrintOut
        
    ActiveWorkbook.Close SaveChanges:=False
    'Pops open the folder with the files when it's done
    Call Shell("explorer.exe" & " " & CurDir & "\Filter Program\1. Filter Condition\Auto Generated", vbNormalFocus)
    
End Sub

'updates a boolean value on the table that prevents this from running on things more than once
Sub UpdateMarkerDateUpdate()

    db.Execute ("UPDATE pg1 " & _
        "SET [pg1].UpdateMarkerDate = TRUE " & _
        "WHERE [pg1].UpdateMarkerDate = FALSE " & _
        "AND [pg1].DateNew IS NOT NULL")

    db.Execute ("UPDATE pg2 " & _
        "SET [pg2].UpdateMarkerDate = TRUE " & _
        "WHERE [pg2].UpdateMarkerDate = FALSE " & _
        "AND [pg2].DateNew IS NOT NULL")

    db.Execute ("UPDATE pg3 " & _
        "SET [pg3].UpdateMarkerDate = TRUE " & _
        "WHERE [pg3].UpdateMarkerDate = FALSE " & _
        "AND [pg3].DateNew IS NOT NULL")

    db.Execute ("UPDATE pg4 " & _
        "SET [pg4].UpdateMarkerDate = TRUE " & _
        "WHERE [pg4].UpdateMarkerDate = FALSE " & _
        "AND [pg4].DateNew IS NOT NULL")

    db.Execute ("UPDATE pg5 " & _
        "SET [pg5].UpdateMarkerDate = TRUE " & _
        "WHERE [pg5].UpdateMarkerDate = FALSE " & _
        "AND [pg5].DateNew IS NOT NULL")

    db.Execute ("UPDATE pg6 " & _
        "SET [pg6].UpdateMarkerDate = TRUE " & _
        "WHERE [pg6].UpdateMarkerDate = FALSE " & _
        "AND [pg6].DateNew IS NOT NULL")

    db.Execute ("UPDATE pg7 " & _
        "SET [pg7].UpdateMarkerDate = TRUE " & _
        "WHERE [pg7].UpdateMarkerDate = FALSE " & _
        "AND [pg7].DateNew IS NOT NULL")
 
End Sub

'I took away people's ability to make new records because the program does that.
'They'll break it if they add a record.
'This just drops them in the last record. 
Private Sub Form_Load()
    
    DoCmd.GoToRecord , , acLast

End Sub

'These chunks just control subform visibility. The people I worked with weren't
'exactly tech savy so I perserved the look of the form in Access which was not particularly
'easy and required 6 subforms.
Private Sub Page_1_Toggle_Click()

    Me!FilterConditionpg2.Visible = False
    Me!FilterConditionpg3.Visible = False
    Me!FilterConditionpg4.Visible = False
    Me!FilterConditionpg5.Visible = False
    Me!FilterConditionpg6.Visible = False
    Me!FilterConditionpg7.Visible = False
    
End Sub


Private Sub Page_2_Toggle_Click()

    If Me!FilterConditionpg2.Visible = True Then
        Me!FilterConditionpg2.Visible = False
    Else
        Me!FilterConditionpg2.Visible = True
        Me!FilterConditionpg3.Visible = False
        Me!FilterConditionpg4.Visible = False
        Me!FilterConditionpg5.Visible = False
        Me!FilterConditionpg6.Visible = False
        Me!FilterConditionpg7.Visible = False
    
    End If

End Sub


Private Sub Page_3_Toggle_Click()

    If Me!FilterConditionpg3.Visible = True Then
        Me!FilterConditionpg3.Visible = False
    Else
        Me!FilterConditionpg3.Visible = True
        Me!FilterConditionpg2.Visible = False
        Me!FilterConditionpg4.Visible = False
        Me!FilterConditionpg5.Visible = False
        Me!FilterConditionpg6.Visible = False
        Me!FilterConditionpg7.Visible = False
    
    End If

End Sub


Private Sub Page_4_Toggle_Click()

    If Me!FilterConditionpg4.Visible = True Then
        Me!FilterConditionpg4.Visible = False
    Else
        Me!FilterConditionpg4.Visible = True
        Me!FilterConditionpg3.Visible = False
        Me!FilterConditionpg2.Visible = False
        Me!FilterConditionpg5.Visible = False
        Me!FilterConditionpg6.Visible = False
        Me!FilterConditionpg7.Visible = False
    
    End If
    
End Sub


Private Sub Page_5_Toggle_Click()

    If Me!FilterConditionpg5.Visible = True Then
        Me!FilterConditionpg5.Visible = False
    Else
        Me!FilterConditionpg5.Visible = True
        Me!FilterConditionpg3.Visible = False
        Me!FilterConditionpg4.Visible = False
        Me!FilterConditionpg2.Visible = False
        Me!FilterConditionpg6.Visible = False
        Me!FilterConditionpg7.Visible = False
    
    End If
    
End Sub


Private Sub Page_6_Toggle_Click()

    If Me!FilterConditionpg6.Visible = True Then
        Me!FilterConditionpg6.Visible = False
    Else
        Me!FilterConditionpg6.Visible = True
        Me!FilterConditionpg3.Visible = False
        Me!FilterConditionpg4.Visible = False
        Me!FilterConditionpg5.Visible = False
        Me!FilterConditionpg2.Visible = False
        Me!FilterConditionpg7.Visible = False
    
    End If
    
End Sub


Private Sub Page_7_Toggle_Click()

    If Me!FilterConditionpg7.Visible = True Then
        Me!FilterConditionpg7.Visible = False
    Else
        Me!FilterConditionpg7.Visible = True
        Me!FilterConditionpg3.Visible = False
        Me!FilterConditionpg4.Visible = False
        Me!FilterConditionpg5.Visible = False
        Me!FilterConditionpg6.Visible = False
        Me!FilterConditionpg2.Visible = False
    
    End If
    
End Sub

'The following 3 subroutines were made because everyone liked to fill out
'their forms differently. So it just prints out 3 different forms that are functionally the same.
'This has 1 week of data and a blank column next to it to write in
Private Sub PrintBlankSecond_Click()
        
    VarDec
    ChangeDec
        
    Dim ConditionPrint As String
    Dim FilterConTab As DAO.Recordset
    Dim KCAPTFC As Excel.Workbook
    Dim TFC As Excel.Worksheet
    Dim ConRow As Integer
    Dim ConCol As Integer
    Dim CleanDate As String
    
    ConditionPrint = TemplateDirectory & "KCAPTFC.xlsx"
    Set FilterConTab = db.OpenRecordset("SELECT [RowDesignation], [OldDate], [CurrentDate], [DateChanged], [DaysUnchanged], [OldCondition], [CurrentCondition] " & _
    "FROM [Filter Condition Data Table]")
    Set KCAPTFC = Excel.Workbooks.Open(ConditionPrint)
    Set TFC = KCAPTFC.Worksheets("FilterDataPrint")
    
    FilterConTab.MoveFirst
    TFC.Cells(5, 12).Value = FilterConTab(2).Value
    TFC.Cells(5, 13).Value = ""
    
    Do While Not FilterConTab.EOF
        TFC.Cells(FilterConTab(0).Value, 8).Value = FilterConTab(3).Value
        TFC.Cells(FilterConTab(0).Value, 9).Value = FilterConTab(4).Value
        TFC.Cells(FilterConTab(0).Value, 12).Value = FilterConTab(6).Value
        TFC.Cells(FilterConTab(0).Value, 13).Value = ""
        FilterConTab.MoveNext
    Loop
    
    FilterConTab.MoveFirst
        
    CleanDate = Replace(CStr(FilterConTab(2).Value), "/", "")
    
    Path = Application.CurrentProject.Path
    ChDir Path
    ChDir ".."
    ChDir ".."
    
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, _
    FileName:=CurDir & "\3.Filter Walk Sheets\3.Second Column Blank\Filter Condition" & CleanDate & ".pdf"
    
    ActiveWorkbook.PrintOut
    
    ActiveWorkbook.Close SaveChanges:=False

End Sub

'Prints out a form with the last 2 weeks of data on it
Private Sub PrintForm_Click()
        
    VarDec
    ChangeDec
        
    Dim ConditionPrint As String
    Dim FilterConTab As DAO.Recordset
    Dim KCAPTFC As Excel.Workbook
    Dim TFC As Excel.Worksheet
    Dim ConRow As Integer
    Dim ConCol As Integer
    Dim CleanDate As String
    
    ConditionPrint = TemplateDirectory & "KCAPTFC.xlsx"
    Set FilterConTab = db.OpenRecordset("SELECT [RowDesignation], [OldDate], [CurrentDate], [DateChanged], [DaysUnchanged], [OldCondition], [CurrentCondition] " & _
    "FROM [Filter Condition Data Table]")
    Set KCAPTFC = Excel.Workbooks.Open(ConditionPrint)
    Set TFC = KCAPTFC.Worksheets("FilterDataPrint")
    
    FilterConTab.MoveFirst
    TFC.Cells(5, 12).Value = FilterConTab(1).Value
    TFC.Cells(5, 13).Value = FilterConTab(2).Value
    
    Do While Not FilterConTab.EOF
        TFC.Cells(FilterConTab(0).Value, 8).Value = FilterConTab(3).Value
        TFC.Cells(FilterConTab(0).Value, 9).Value = FilterConTab(4).Value
        TFC.Cells(FilterConTab(0).Value, 12).Value = FilterConTab(5).Value
        TFC.Cells(FilterConTab(0).Value, 13).Value = FilterConTab(6).Value
        FilterConTab.MoveNext
    Loop
    
    FilterConTab.MoveFirst
    
    CleanDate = Replace(CStr(FilterConTab(2).Value), "/", "")
    
    Path = Application.CurrentProject.Path
    ChDir Path
    ChDir ".."
    ChDir ".."
    
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, _
    FileName:=CurDir & "\Filter Program\3.Filter Walk Sheets\2.Double Column\Filter Condition" & CleanDate & ".pdf"
    
    ActiveWorkbook.PrintOut
    
    ActiveWorkbook.Close SaveChanges:=False

End Sub

'This prints out a form with 1 column containing last week's data. no notes column.
Private Sub PrintSingle_Click()

    VarDec
    ChangeDec
        
    Dim ConditionPrint As String
    Dim FilterConTab As DAO.Recordset
    Dim KCAPTFCS As Excel.Workbook
    Dim TFCS As Excel.Worksheet
    Dim ConRow As Integer
    Dim ConCol As Integer
    Dim CleanDate As String
    
    ConditionPrint = TemplateDirectory & "KCAPTFCS.xlsx"
    Set FilterConTab = db.OpenRecordset("SELECT [RowDesignation], [CurrentDate], [DateChanged], [DaysUnchanged], [CurrentCondition] " & _
    "FROM [Filter Condition Data Table]")
    Set KCAPTFCS = Excel.Workbooks.Open(ConditionPrint)
    Set TFCS = KCAPTFCS.Worksheets("FilterDataPrint")
    
    FilterConTab.MoveFirst
    TFCS.Cells(5, 12).Value = FilterConTab(1).Value
    
    Do While Not FilterConTab.EOF
        TFCS.Cells(FilterConTab(0).Value, 8).Value = FilterConTab(2).Value
        TFCS.Cells(FilterConTab(0).Value, 9).Value = FilterConTab(3).Value
        TFCS.Cells(FilterConTab(0).Value, 12).Value = FilterConTab(4).Value
        FilterConTab.MoveNext
    Loop
    
    FilterConTab.MoveFirst
    
    CleanDate = Replace(CStr(FilterConTab(1).Value), "/", "")
    
    Path = Application.CurrentProject.Path
    ChDir Path
    ChDir ".."
    ChDir ".."
    
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, _
    FileName:=CurDir & "\Filter Program\3.Filter Walk Sheets\1.Single Column\Filter Condition" & CleanDate & ".pdf"
       
    'ActiveWorkbook.PrintOut
    
    ActiveWorkbook.Close SaveChanges:=False

End Sub

'This was an experiment to try to make it generate the report faster
'using power query and an excel file. I never finished it as it was more a nice to have

'Private Sub QuickCOR_Click()
'
'    ChangeDec
''    ScheduledChangeOut
''    RedChangeOut
' '   YellowChangeOut
' '   ChangeOutExport
'
'    Dim ChangeRequest As String
'    Dim FilterChangeTable As DAO.Recordset
'    Dim KCAPTFQuickStatus As Excel.Workbook
'    Dim COR As Excel.Worksheet
'    Dim CRow As Integer
'    Dim CCol As Integer
'    Dim ColTitleRowF As String
'    Dim ColTitleRowS As String
'    Dim ColTitleRowL As String
'    Dim SecTitleRowF As String
'    Dim SecTitleRowS As String
'    Dim SecTitleRowL As String
'    Dim SaveDatePDF As DAO.Recordset
'    Dim CleanDate As String
'
'    DoCmd.TransferDatabase acExport, "Microsoft Access", TemplateDirectory & "KCAPTFQuickStatus.accdb", acTable, "Filter Change Out Request", "Filter Change Out Request"
'
'    ChangeRequest = TemplateDirectory & "KCAPTFQuickStatus.xlsx"
'    Set FilterChangeTable = db.OpenRecordset("SELECT * FROM [Filter Change Out Request]")
'    Set KCAPTFQuickStatus = Excel.Workbooks.Open(ChangeRequest, 3)
'    Set COR = KCAPTFQuickStatus.Worksheets("Change Out Request")
'    CRow = 1
'    CCol = 1
'    ColTitleRowF = "A2:M2"
'    ColTitleRowS = "A" & (3 + SN) & ":M" & (3 + SN)
'    ColTitleRowL = "A" & (4 + SN + RN) & ":M" & (4 + SN + RN)
'    SecTitleRowF = "A1:M1"
'    SecTitleRowS = "A" & (2 + SN) & ":M" & (2 + SN)
'    SecTitleRowL = "A" & (3 + SN + RN) & ":M" & (3 + SN + RN)
'
'    COR.Range("A1:M150").Value = ""
'    COR.Range("A1:M150").Font.FontStyle = "Regular"
'    COR.Range("A1:M150").Font.Underline = xlUnderlineStyleNone
'    COR.Range("A1:M150").Font.Size = 14
'    COR.Range("A1:M150").Font.Color = 0
'    COR.Range("A1:M150").Borders.LineStyle = xlNone
'    COR.Range("A1:M150").Interior.ColorIndex = 0
'    COR.Range("A1:M150").HorizontalAlignment = xlHAlignCenter
'    COR.Range("A1:M150").UnMerge
'
'    If BGRED > 0 Then
'        COR.Range("A" & (4 + SN) & ":M" & (3 + SN + BGRED)).Interior.Color = 255
'    End If
'
'    If BGYREDTotal > 0 Then
'        COR.Range("A" & (4 + SN + BGRED) & ":M" & (3 + SN + BGYREDEnd)).Interior.Color = 65535
'    End If
'
'    COR.Range("k3:l" & (1 + SN)).Font.Color = 255
'    COR.Range("A1:M" & (SN + RN + YN + 3)).Borders.LineStyle = xlDash
'
'
'    COR.Range(SecTitleRowF).Font.FontStyle = "Bold"
'    COR.Range(SecTitleRowS).Font.FontStyle = "Bold"
'    COR.Range(SecTitleRowL).Font.FontStyle = "Bold"
'    COR.Range(SecTitleRowF).Borders.LineStyle = xlNone
'    COR.Range(SecTitleRowS).Borders.LineStyle = xlNone
'    COR.Range(SecTitleRowL).Borders.LineStyle = xlNone
'    COR.Range(SecTitleRowF).Font.Size = 16
'    COR.Range(SecTitleRowS).Font.Size = 16
'    COR.Range(SecTitleRowL).Font.Size = 16
'    COR.Range(SecTitleRowF).Font.Color = 15773696
'    COR.Range(SecTitleRowS).Font.Color = 15773696
'    COR.Range(SecTitleRowL).Font.Color = 15773696
'    COR.Range(SecTitleRowF).Merge
'    COR.Range(SecTitleRowS).Merge
'    COR.Range(SecTitleRowL).Merge
'    COR.Range(SecTitleRowF).HorizontalAlignment = xlLeft
'    COR.Range(SecTitleRowS).HorizontalAlignment = xlLeft
'    COR.Range(SecTitleRowL).HorizontalAlignment = xlLeft
'
'    COR.Range(ColTitleRowF).Font.FontStyle = "Bold Italic"
'    COR.Range(ColTitleRowS).Font.FontStyle = "Bold Italic"
'    COR.Range(ColTitleRowL).Font.FontStyle = "Bold Italic"
'    COR.Range(ColTitleRowF).Font.Underline = xlUnderlineStyleSingle
'    COR.Range(ColTitleRowS).Font.Underline = xlUnderlineStyleSingle
'    COR.Range(ColTitleRowL).Font.Underline = xlUnderlineStyleSingle
'    COR.Range(ColTitleRowF).Font.Size = 16
'    COR.Range(ColTitleRowS).Font.Size = 16
'    COR.Range(ColTitleRowL).Font.Size = 16
'    COR.Range(ColTitleRowF).Font.Color = 6299648
'    COR.Range(ColTitleRowS).Font.Color = 6299648
'    COR.Range(ColTitleRowL).Font.Color = 6299648
'    COR.Range(ColTitleRowF).Borders.LineStyle = xlContinuous
'    COR.Range(ColTitleRowS).Borders.LineStyle = xlContinuous
'    COR.Range(ColTitleRowL).Borders.LineStyle = xlContinuous
'    COR.Range(ColTitleRowF).Interior.Color = 13434879
'    COR.Range(ColTitleRowS).Interior.Color = 13434879
'    COR.Range(ColTitleRowL).Interior.Color = 13434879
'
'    Set SaveDatePDF = db.OpenRecordset("SELECT [CurrentDate] FROM [Filter Condition Data Table]")
'
'    SaveDatePDF.MoveFirst
'
'    CleanDate = Replace(CStr(SaveDatePDF(0).Value), "/", "")
'
'
'    Path = Application.CurrentProject.Path
'    ChDir Path
'    ChDir ".."
'    ChDir ".."
'
'    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, _
'    FileName:=CurDir & "\Filter Program\1. Filter Condition\Auto Generated\Filter Change Request" & CleanDate & ".pdf"
'
'    ActiveWorkbook.SaveAs FileName:=CurDir & "\Filter Program\1. Filter Condition\Auto Generated\Filter Change Request" & CleanDate & ".xlsx"
'
'    'ActiveWorkbook.PrintOut
'
'    ActiveWorkbook.Close SaveChanges:=False
'
'    Call Shell("explorer.exe" & " " & CurDir & "\Filter Program\1. Filter Condition\Auto Generated", vbNormalFocus)
'
'End Sub

'Automatically opens subforms when you leave the last data entry block on the form
Private Sub Textpg1FinalRecord_Exit(Cancel As Integer)

    Me!FilterConditionpg2.Visible = True
    Me!FilterConditionpg2.SetFocus
    
End Sub






