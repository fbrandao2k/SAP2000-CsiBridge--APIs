' routine a cable stayed bridge (ponte estaiada)
' 2D static analysis of the cable-stayed bridge
' it is assumed: 2D, and cables
' Template started with Default Units (kN, m, C)
' Default Materials Europe (saved as default template in SAP2000)
' Before starting, open another instance of SAP2000, and delete the folder if the project was already created
' Make sure the you have access to the folder in which the file is being saved, and the file is already created

' Author Felipe da Silva Brandao

Sub POF_4()
    
    Dim ret As Long ' ret = 0 when the command works, and =1 when it does not work

    'set the following flag to True to attach to an existing instance of the program
    'otherwise a new instance of the program will be started

    '*********************************************************************************************
    'INITIALIZE THE MODEL
        
    Dim AttachToInstance As Boolean
    AttachToInstance = False
 
    'set the following flag to True to manually specify the path to SAP2000.exe
    'this allows for a connection to a version of SAP2000 other than the latest installation
    'otherwise the latest installed version of SAP2000 will be launched

    Dim SpecifyPath As Boolean
    SpecifyPath = False
 
    'if the above flag is set to True, specify the path to SAP2000 below

    Dim ProgramPath As String
    'ProgramPath = "C:\Program Files\Computers and Structures\SAP2000 22\SAP2000.exe"

    'full path to the model
    'set it to the desired path of your model

    Dim ModelDirectory As String
    ModelDirectory = "C:\Users\FelipeSBrandao\SAP2000"
    If Len(Dir(ModelDirectory, vbDirectory)) = 0 Then
        MkDir ModelDirectory
    End If

    Dim ModelName As String
    ModelName = "SAP_1-POF.sdb"
    Dim ModelPath As String

    ModelPath = ModelDirectory & Application.PathSeparator & ModelName
    
    'create API helper object

    Dim myHelper As cHelper
    Set myHelper = New Helper

    'dimension the SapObject as cOAPI type
    Dim mySapObject As cOAPI
    Set mySapObject = Nothing

    If AttachToInstance Then
      'attach to a running instance of SAP2000
      'get the active SapObject
      Set mySapObject = myHelper.GetObject("CSI.SAP2000.API.SapObject")
    Else

      If SpecifyPath Then
        'create an instance of the SapObject from the specified path
        Set mySapObject = myHelper.CreateObject(ProgramPath)
      Else

        'create an instance of the SapObject from the latest installed SAP2000
        Set mySapObject = myHelper.CreateObjectProgID("CSI.SAP2000.API.SapObject")
      End If

    '*********************************************************************************************
    'START SAP2000
      mySapObject.ApplicationStart
    End If

    'Get a reference to cSapModel to access all OAPI classes and functions

    Dim mySapModel As cSapModel
    Set mySapModel = mySapObject.SapModel

    'initialize model
    ret = mySapModel.InitializeNewModel

    '*********************************************************************************************
    ' CREATE A MODEL
    ret = mySapModel.File.NewBlank
    
    'switch to kN_m_C units
    ret = mySapModel.SetPresentUnits(eUnits_kN_m_C)

    '*********************************************************************************************
    ' DEFINE MATERIAL CONCRETE PROPERTY
    
    Dim Name As String
    Dim E As Long
    
    E = Worksheets("Secoes").Range("M11").Value
    
    'creation of the material concrete (for the deck)
    ret = mySapModel.PropMaterial.SetMaterial("CONCRETE", eMatType_Concrete)
    'assign isotropic mechanical properties to material (name, E, poisson, Coeff of thermal Expansion, [temperatura (optional)] )
    ret = mySapModel.PropMaterial.SetMPIsotropic("CONCRETE", E, 0.2, 0.00001)
    'assign material property weight per unit volume
    ret = mySapModel.PropMaterial.SetWeightAndMass("CONCRETE", 1, Worksheets("Secoes").Range("M12").Value)  ' Kn /M3
    ret = mySapModel.PropMaterial.AddQuick(Name, eMatType.eMatType_Rebar, , , , , eMatTypeRebar.eMatTypeRebar_ASTM_A706)
    
    '*********************************************************************************************
    ' DEFINE CABLE CONCRETE PROPERTY
    
    'creation of the material tendon (for the cables)
    ret = mySapModel.PropMaterial.SetMaterial("ESTAI", eMatType_Tendon)
    'assign isotropic mechanical properties to material (name, E, poisson, Coeff of thermal Expansion, [temperatura (optional)] )
    ret = mySapModel.PropMaterial.SetMPIsotropic("ESTAI", 195000000, 0, 0.00001) ' 195 GPa
    ret = mySapModel.PropMaterial.SetWeightAndMass("ESTAI", 1, 76.97) ' Kn /M3 'assign material property weight per unit volume

    'define User Defined frame section property - Deck of Ponte Estaiada Octavio Frias
    'ret = mySapModel.PropFrame.SetRectangle("R1", "CONC", 12, 12)
    'Name = "Aco_CA50"
    'add ASTM A706 rebar material
    ret = mySapModel.PropMaterial.AddQuick(Name, eMatType.eMatType_Rebar, , , , , eMatTypeRebar.eMatTypeRebar_ASTM_A706)
      
    '*********************************************************************************************
    ' ADD FRAME SECTIONS
    
    Dim i As Integer
    Dim j As Integer
    Dim NumberPoints As Long
    Dim X() As Double
    Dim Y() As Double
    Dim Ycg As Double ' centroid of the deck_POF main section
    Dim Radius() As Double
    Dim nSections As Integer
    Dim sectionName As String
    
    Ycg = Worksheets("Secoes").Range("F12").Value ' centroid of the deck_POF main section
    nSections = Worksheets("Secoes").Range("L6").Value ' number of different frame sections = 6
    
    j = 1 ' counts the number of sections that are being created
    While j <= nSections
        sectionName = Worksheets("Secoes").Range("F" & CStr((20 * (j - 1)) + 2)).Value
        ret = mySapModel.PropFrame.SetSDSection(sectionName, "CONCRETE") '(name, material) 'add new section designer frame section property
        'add polygon shape to new property
        NumberPoints = Worksheets("Secoes").Range("F" & CStr((20 * (j - 1)) + 2) + 12).Value
        ReDim X(NumberPoints - 1)
        ReDim Y(NumberPoints - 1)
        ReDim Radius(NumberPoints - 1)
    
        For i = 1 To NumberPoints
            X(i - 1) = Worksheets("Secoes").Range("A" & CStr((20 * (j - 1)) + i + 1)).Value
            Y(i - 1) = Worksheets("Secoes").Range("B" & CStr((20 * (j - 1)) + i + 1)).Value
            Radius(i - 1) = Worksheets("Secoes").Range("C" & CStr((20 * (j - 1)) + i + 1)).Value
        Next i
        ret = mySapModel.PropFrame.SDShape.SetPolygon(sectionName, "ShapeDeck" & CStr(j), "CONCRETE", "Default", NumberPoints, X, Y, Radius, -1, True, Name)
        j = j + 1
    Wend
 
    
    '*********************************************************************************************
    ' CREATION OF NON-PRISMATIC SECTIONS
    
    Dim StartSec() As String
    Dim EndSec() As String
    Dim MyLength() As Double
    Dim MyType() As Long
    Dim EI33() As Long
    Dim EI22() As Long
    Dim nSegments As Integer
    
    nSections = Worksheets("Secoes_VAR").Range("A14").Value ' number of different frame sections with variations= 6
    
    For i = 1 To nSections
        nSegments = Worksheets("Secoes_VAR").Range("B" & CStr(i + 1)).Value
        ReDim StartSec(nSegments - 1)
        ReDim EndSec(nSegments - 1)
        ReDim MyLength(nSegments - 1)
        ReDim MyType(nSegments - 1)
        ReDim EI33(nSegments - 1)
        ReDim EI22(nSegments - 1)
        StartSec(0) = Worksheets("Secoes_VAR").Range("C" & CStr(i + 1)).Value
        EndSec(0) = Worksheets("Secoes_VAR").Range("D" & CStr(i + 1)).Value
        MyLength(0) = Worksheets("Secoes_VAR").Range("E" & CStr(i + 1)).Value
        MyType(0) = Worksheets("Secoes_VAR").Range("F" & CStr(i + 1)).Value  ' 1= relative lenght, 2 = absolute lenght
        EI33(0) = Worksheets("Secoes_VAR").Range("G" & CStr(i + 1)).Value  ' 1 = linear variation of EI33
        EI22(0) = Worksheets("Secoes_VAR").Range("H" & CStr(i + 1)).Value
        
        If nSegments = 2 Then
            StartSec(1) = Worksheets("Secoes_VAR").Range("I" & CStr(i + 1)).Value
            EndSec(1) = Worksheets("Secoes_VAR").Range("J" & CStr(i + 1)).Value
            MyLength(1) = Worksheets("Secoes_VAR").Range("K" & CStr(i + 1)).Value
            MyType(1) = Worksheets("Secoes_VAR").Range("L" & CStr(i + 1)).Value  ' 1= relative lenght, 2 = absolute lenght
            EI33(1) = Worksheets("Secoes_VAR").Range("M" & CStr(i + 1)).Value  ' 1 = linear variation of EI33
            EI22(1) = Worksheets("Secoes_VAR").Range("N" & CStr(i + 1)).Value
        End If
        
        ret = mySapModel.PropFrame.SetNonPrismatic(Worksheets("Secoes_VAR").Range("A" & CStr(i + 1)).Value, nSegments, StartSec, EndSec, MyLength, MyType, EI33, EI22)
    Next i

    '*********************************************************************************************
    ' ADD FRAMES OBJECTS BY COORDINATES
    ' add frames, and create the group nodes_DECK that is assigned to the points
    
    ret = mySapModel.GroupDef.SetGroup("nodes_DECK") 'define new group called nodes_DECK for the points of the deck
    
    Dim nFrames As Double
    Dim FrameName() As String
    Dim Offset1() As Double 'offsets of the centroid of insertion point
    Dim Offset2() As Double
    ReDim Offset1(2)
    ReDim Offset2(2)
    
    Offset1(2) = -Ycg ' when insertion point = bottom center
    Offset2(2) = -Ycg
    
    nFrames = Worksheets("Frames").Range("M2").Value
    ReDim FrameName(nFrames - 1)

    With Worksheets("Frames") ' frame from 1 to 19
        For i = 1 To 18 'nFrames                                 'xi                           'yi                             'zi                              'xf                          'yf                          'zf                              'Empty Name       'section name
            ret = mySapModel.FrameObj.AddByCoord(.Range("B" & CStr(i + 1)).Value, .Range("C" & CStr(i + 1)).Value, .Range("D" & CStr(i + 1)).Value, .Range("E" & CStr(i + 1)).Value, .Range("F" & CStr(i + 1)).Value, .Range("G" & CStr(i + 1)).Value, FrameName(i - 1), "deck_POF")
            ret = mySapModel.PointObj.SetGroupAssign(CStr(i), "nodes_DECK") 'add point object to group
        Next i
    
        i = 19
        ret = mySapModel.FrameObj.AddByCoord(.Range("B" & CStr(i + 1)).Value, .Range("C" & CStr(i + 1)).Value, .Range("D" & CStr(i + 1)).Value, .Range("E" & CStr(i + 1)).Value, .Range("F" & CStr(i + 1)).Value, .Range("G" & CStr(i + 1)).Value, FrameName(i - 1), "deck_POF_VAR1")
        ret = mySapModel.FrameObj.SetInsertionPoint(CStr(i), 2, False, True, Offset1, Offset2, "Global") 'name ' 2=bottom center 'False = No mirror ' True = Stiff Transf
        ret = mySapModel.PointObj.SetGroupAssign(CStr(i), "nodes_DECK") 'add point object to group
        i = 20
        ret = mySapModel.FrameObj.AddByCoord(.Range("B" & CStr(i + 1)).Value, .Range("C" & CStr(i + 1)).Value, .Range("D" & CStr(i + 1)).Value, .Range("E" & CStr(i + 1)).Value, .Range("F" & CStr(i + 1)).Value, .Range("G" & CStr(i + 1)).Value, FrameName(i - 1), "deck_POF_VAR2")
        ret = mySapModel.FrameObj.SetInsertionPoint(CStr(i), 2, False, True, Offset1, Offset2, "Global") 'name ' 2=bottom center 'False = No mirror ' True = Stiff Transf
        ret = mySapModel.PointObj.SetGroupAssign(CStr(i), "nodes_DECK") 'add point object to group
        i = 21
        ret = mySapModel.FrameObj.AddByCoord(.Range("B" & CStr(i + 1)).Value, .Range("C" & CStr(i + 1)).Value, .Range("D" & CStr(i + 1)).Value, .Range("E" & CStr(i + 1)).Value, .Range("F" & CStr(i + 1)).Value, .Range("G" & CStr(i + 1)).Value, FrameName(i - 1), "deck_POF_VAR3")
        ret = mySapModel.FrameObj.SetInsertionPoint(CStr(i), 2, False, True, Offset1, Offset2, "Global") 'name ' 2=bottom center 'False = No mirror ' True = Stiff Transf
        ret = mySapModel.PointObj.SetGroupAssign(CStr(i), "nodes_DECK") 'add point object to group
        i = 22
        ret = mySapModel.FrameObj.AddByCoord(.Range("B" & CStr(i + 1)).Value, .Range("C" & CStr(i + 1)).Value, .Range("D" & CStr(i + 1)).Value, .Range("E" & CStr(i + 1)).Value, .Range("F" & CStr(i + 1)).Value, .Range("G" & CStr(i + 1)).Value, FrameName(i - 1), "deck_POF_Mastro")
        ret = mySapModel.FrameObj.SetInsertionPoint(CStr(i), 2, False, True, Offset1, Offset2, "Global") 'name ' 2=bottom center 'False = No mirror ' True = Stiff Transf
        ret = mySapModel.PointObj.SetGroupAssign(CStr(i), "nodes_DECK") 'add point object to group
        i = 23
        ret = mySapModel.FrameObj.AddByCoord(.Range("B" & CStr(i + 1)).Value, .Range("C" & CStr(i + 1)).Value, .Range("D" & CStr(i + 1)).Value, .Range("E" & CStr(i + 1)).Value, .Range("F" & CStr(i + 1)).Value, .Range("G" & CStr(i + 1)).Value, FrameName(i - 1), "deck_POF_Mastro")
        ret = mySapModel.FrameObj.SetInsertionPoint(CStr(i), 2, False, True, Offset1, Offset2, "Global") 'name ' 2=bottom center 'False = No mirror ' True = Stiff Transf
        ret = mySapModel.PointObj.SetGroupAssign(CStr(i), "nodes_DECK") 'add point object to group
        i = 24
        ret = mySapModel.FrameObj.AddByCoord(.Range("B" & CStr(i + 1)).Value, .Range("C" & CStr(i + 1)).Value, .Range("D" & CStr(i + 1)).Value, .Range("E" & CStr(i + 1)).Value, .Range("F" & CStr(i + 1)).Value, .Range("G" & CStr(i + 1)).Value, FrameName(i - 1), "deck_POF_VAR4")
        ret = mySapModel.FrameObj.SetInsertionPoint(CStr(i), 2, False, True, Offset1, Offset2, "Global") 'name ' 2=bottom center 'False = No mirror ' True = Stiff Transf
        ret = mySapModel.PointObj.SetGroupAssign(CStr(i), "nodes_DECK") 'add point object to group
        i = 25
        ret = mySapModel.FrameObj.AddByCoord(.Range("B" & CStr(i + 1)).Value, .Range("C" & CStr(i + 1)).Value, .Range("D" & CStr(i + 1)).Value, .Range("E" & CStr(i + 1)).Value, .Range("F" & CStr(i + 1)).Value, .Range("G" & CStr(i + 1)).Value, FrameName(i - 1), "deck_POF_VAR5")
        ret = mySapModel.FrameObj.SetInsertionPoint(CStr(i), 2, False, True, Offset1, Offset2, "Global") 'name ' 2=bottom center 'False = No mirror ' True = Stiff Transf
        ret = mySapModel.PointObj.SetGroupAssign(CStr(i), "nodes_DECK") 'add point object to group
        i = 26
        ret = mySapModel.FrameObj.AddByCoord(.Range("B" & CStr(i + 1)).Value, .Range("C" & CStr(i + 1)).Value, .Range("D" & CStr(i + 1)).Value, .Range("E" & CStr(i + 1)).Value, .Range("F" & CStr(i + 1)).Value, .Range("G" & CStr(i + 1)).Value, FrameName(i - 1), "deck_POF_VAR6")
        ret = mySapModel.FrameObj.SetInsertionPoint(CStr(i), 2, False, True, Offset1, Offset2, "Global") 'name ' 2=bottom center 'False = No mirror ' True = Stiff Transf
        ret = mySapModel.PointObj.SetGroupAssign(CStr(i), "nodes_DECK") 'add point object to group
        
        For i = 27 To 44 'last deck frame 'nFrames                   'xi                           'yi                             'zi                              'xf                          'yf                          'zf                              'Empty Name       'section name
            ret = mySapModel.FrameObj.AddByCoord(.Range("B" & CStr(i + 1)).Value, .Range("C" & CStr(i + 1)).Value, .Range("D" & CStr(i + 1)).Value, .Range("E" & CStr(i + 1)).Value, .Range("F" & CStr(i + 1)).Value, .Range("G" & CStr(i + 1)).Value, FrameName(i - 1), "deck_POF")
            ret = mySapModel.PointObj.SetGroupAssign(CStr(i), "nodes_DECK") 'add point object to group
        Next i
        ret = mySapModel.PointObj.SetGroupAssign(CStr(i), "nodes_DECK") 'add last point object to group
    End With
    'ret = mySapModel.FrameObj.AddByCoord(0, 0, 0, 0, 0, 10, FrameName(0), "R1", "1")
    'ret = mySapModel.FrameObj.AddByCoord(0, 0, 10, 8, 0, 16, FrameName(1), "R1", "2")
    'ret = mySapModel.FrameObj.AddByCoord(-4, 0, 10, 0, 0, 10, FrameName(2), "R1", "3")
    
    '*********************************************************************************************
    ' ADD LINK OBJECTS BY COORDINATES
    ' add links joining cables from left and right to the central frame deck
    ' it also creates the groups nodes_External and nodes_Internal to assign to these new points
    
    Dim nCables As Integer
    Dim xi As Double
    Dim yi As Double
    Dim zi As Double
    Dim xf As Double
    Dim yf As Double
    Dim zf As Double
    nCables = 36 * 2
    
    'add link property
    Dim DOF() As Boolean
    Dim Fixed() As Boolean
    ReDim DOF(5)
    ReDim Fixed(5)
    Dim Ke() As Double
    Dim Ce() As Double
    ReDim Ke(5)
    ReDim Ce(5)
   'ReDim Ke(5)
   'ReDim Ce(5)
    DOF(0) = True
    Fixed(0) = True
    DOF(1) = True
    Fixed(1) = True
    DOF(2) = True
    Fixed(2) = True
    'Ke(0) = 0
    'Ke(1) = 0
    'Ke(2) = 0
    'Ce(0) = 0
    'Ce(1) = 0
    'Ce(2) = 0
    ret = mySapModel.PropLink.SetLinear("RigdLink", DOF, Fixed, Ke, Ce, 0, 0)
      
    'add link object by coordinates
    For i = 1 To nCables
        xi = Worksheets("Links").Range("B" & CStr(i + 1)).Value
        yi = Worksheets("Links").Range("C" & CStr(i + 1)).Value
        zi = Worksheets("Links").Range("D" & CStr(i + 1)).Value
        xf = Worksheets("Links").Range("E" & CStr(i + 1)).Value
        yf = Worksheets("Links").Range("F" & CStr(i + 1)).Value
        zf = Worksheets("Links").Range("G" & CStr(i + 1)).Value
        ret = mySapModel.LinkObj.AddByCoord(xi, yi, zi, xf, yf, zf, Name, False) 'True = 'one joint link ; False = two joints links
    Next i

    '*********************************************************************************************
    ' ADD CABLES SECTIONS
    ' ADD STAYS SECTIONS
    
    For i = 1 To nCables '                               'name                                  'material                          'area
        ret = mySapModel.PropCable.SetProp(Worksheets("Stays").Range("M" & CStr(i + 1)).Value, "ESTAI", Worksheets("Stays").Range("Q" & CStr(i + 1)).Value)
    Next i

    '*********************************************************************************************
    ' ADD CABLES OBJECTS BY COORDINATES
    ' ADD STAYS OBJECTS BY COORDINATES
    
    Dim cableName As String
       
    For i = 1 To nCables
        xi = Worksheets("Stays").Range("C" & CStr(i + 1)).Value
        yi = Worksheets("Stays").Range("D" & CStr(i + 1)).Value
        zi = Worksheets("Stays").Range("E" & CStr(i + 1)).Value
        xf = Worksheets("Stays").Range("I" & CStr(i + 1)).Value
        yf = Worksheets("Stays").Range("J" & CStr(i + 1)).Value
        zf = Worksheets("Stays").Range("K" & CStr(i + 1)).Value
        cableName = Worksheets("Stays").Range("M" & CStr(i + 1)).Value
        ret = mySapModel.CableObj.AddByCoord(xi, yi, zi, xf, yf, zf, Name, cableName, cableName) ' cableName here is the same name of the section of the cable
    Next i


    '*********************************************************************************************
    ' SET ACTIVE DOFs
    'set model degrees of freedom
    ''DOF(0) = UX; DOF(1) = UY; DOF(2) = UZ; DOF(3) = RX; DOF(4) = RY; DOF(5) = RZ
    ReDim DOF(5)
    DOF(0) = True
    DOF(1) = True
    DOF(2) = True
    DOF(3) = True
    DOF(4) = True
    DOF(5) = True
    
    ret = mySapModel.Analyze.SetActiveDOF(DOF)

    '*********************************************************************************************
    ' ASSIGN POINT RESTRAINTS
    ' choose the frame/point of the start, of the middle, and the end for the restraints
    ' It also adds cable points restraints

    Dim PointName As String
    Dim Restraint_Start() As Boolean
    Dim Restraint_Middle() As Boolean
    Dim Restraint_End() As Boolean
    Dim Restraint_Cable() As Boolean
    ReDim Restraint_Start(5)
    ReDim Restraint_Middle(5)
    ReDim Restraint_End(5)
    ReDim Restraint_Cable(5)

    For i = 1 To 6 ' number of DOFs
        Restraint_Start(i - 1) = Worksheets("Restraints").Range("B" & CStr(i + 1)).Value
        Restraint_Middle(i - 1) = Worksheets("Restraints").Range("C" & CStr(i + 1)).Value
        Restraint_End(i - 1) = Worksheets("Restraints").Range("D" & CStr(i + 1)).Value
        Restraint_Cable(i - 1) = Worksheets("Restraints").Range("E" & CStr(i + 1)).Value
    Next i

    ' first node retraint
    PointName = Worksheets("Restraints").Range("B1").Value
    ret = mySapModel.PointObj.SetRestraint(PointName, Restraint_Start)
    ' middle node retraint = Z
    PointName = Worksheets("Restraints").Range("C1").Value
    ret = mySapModel.PointObj.SetRestraint(PointName, Restraint_Middle)
    ' end node retraint = Z
    PointName = Worksheets("Restraints").Range("D1").Value
    ret = mySapModel.PointObj.SetRestraint(PointName, Restraint_End)
    
    'assign restraint to the end of the cables (in the Mastro)
    For i = 1 To nCables / 2 ' in this example the top nodes repeats two cables, for example m2e-18 and m2i-18 have the same top point node
        PointName = Worksheets("Stays").Range("G" & CStr(i * 2)).Value
        ret = mySapModel.PointObj.SetRestraint(PointName, Restraint_Cable)
    Next i
   
    'refresh view, update (initialize) zoom
    'ret = mySapModel.View.RefreshView(0, False)
    ret = mySapModel.View.RefreshView
  
    '*********************************************************************************************
    ' ADD LOAD PATTERNS
    
    ' DEAD load pattern is  automatically    'Name           'type   'self weight enable
    'ret = mySapModel.LoadPatterns.Add("Peso_Proprio", LTYPE_DEAD, 1)
    
    ret = mySapModel.LoadPatterns.Add("PROT", eLoadPatternType_Other)
    ret = mySapModel.LoadPatterns.Add("PAVIMENTACAO", eLoadPatternType_Other)
    ret = mySapModel.LoadPatterns.Add("GUARDARODAS", eLoadPatternType_Other)
    ret = mySapModel.LoadPatterns.Add("DUTOPLACATUBOS", eLoadPatternType_Other)
    ret = mySapModel.LoadPatterns.Add("ENCHIMENTOS", eLoadPatternType_Other)
    
    '*********************************************************************************************
    ' ADD LOAD CASE DEFINITIONS
    
    'initialize stage definitions
    Dim MyDuration() As Long
    Dim MyOutput() As Boolean
    Dim MyOutputName() As String
    Dim MyComment() As String
    Dim nameCase As String
    
    Dim nStageDefinitions As Integer
    Dim nCases As Integer   'number of load cases
  
    nStageDefinitions = 1 'definitions for each case
    nCases = Worksheets("Loads_Cases").Range("R1").Value 'in this example, 3 cases="DEAD+OTHERS+PROT", "DEAD+OTHERS+PROT+ENC", and "DEAD+PAV+PROT"
    
    ReDim MyDuration(nStageDefinitions - 1)
    ReDim MyOutput(nStageDefinitions - 1)
    ReDim MyOutputName(nStageDefinitions - 1)
    ReDim MyComment(nStageDefinitions - 1)
 
    For i = 1 To nCases ' here it is considered all cases have only one staged definition
        nameCase = Worksheets("Loads_Cases").Range("M" & CStr(i + 2)).Value
        ret = mySapModel.LoadCases.StaticNonlinearStaged.SetCase(nameCase)
        MyDuration(0) = Worksheets("Loads_Cases").Range("N" & CStr(i + 2)).Value ' there are not iterations in this example!
        MyOutput(0) = True
        MyOutputName(0) = Worksheets("Loads_Cases").Range("P" & CStr(i + 2)).Value
        MyComment(0) = Worksheets("Loads_Cases").Range("Q" & CStr(i + 2)).Value
        'creates a Static Non Linear Stages Definition               'load case name    'number of stages=1,2,3... 'duration time
        ret = mySapModel.LoadCases.StaticNonlinearStaged.SetStageDefinitions_1(nameCase, 1, MyDuration, MyOutput, MyOutputName, MyComment)
        ret = mySapModel.LoadCases.StaticNonlinearStaged.SetTargetForceParameters(nameCase, 0.001, 10000000, 1, False) ' set target force parameters, such as TolConvF, MaxIter, AccelFact, NoStop (continues analysis if no convergence)
    Next i
     
    
    '*********************************************************************************************
    ' ADD LOAD CASE STAGE DATA INITIALIZE
    
    'set stage data
    Dim MyOperation() As Long
    Dim MyObjectType() As String
    Dim MyObjectName() As String
    Dim MyAge() As Long
    Dim MyMyType() As String
    Dim MyMyName() As String
    Dim MySF() As Double
    Dim nOperations As Integer
    Dim iStart As Integer ' indicates in which row the range starts the counting
    
    '*********************************************************************************************
'    ' LOAD CASE STAGE DATA 1 (DEAD + OTHERS (PAV, GUARDA RODAS, DUTOS) + PROTENSAO) ********
'
'    iStart = 1 ' indicates the start of DEAD+OTHERS+PROT load case (of the header)
'    nOperations = Worksheets("Loads_Cases").Range("K" & CStr(iStart)).Value 'operations inside the stage definition, such as Add Structure, Load Objects,...
'
'    ReDim MyOperation(nOperations - 1)
'    ReDim MyObjectType(nOperations - 1)
'    ReDim MyObjectName(nOperations - 1)
'    ReDim MyAge(nOperations - 1)
'    ReDim MyMyType(nOperations - 1)
'    ReDim MyMyName(nOperations - 1)
'    ReDim MySF(nOperations - 1)
'
'    For i = 1 To nOperations
'        MyOperation(i - 1) = Worksheets("Loads_Cases").Range("A" & CStr(i + iStart + 1)).Value
'        MyObjectType(i - 1) = Worksheets("Loads_Cases").Range("B" & CStr(i + iStart + 1)).Value
'        MyObjectName(i - 1) = Worksheets("Loads_Cases").Range("C" & CStr(i + iStart + 1)).Value
'        MyAge(i - 1) = Worksheets("Loads_Cases").Range("D" & CStr(i + iStart + 1)).Value
'        MyMyType(i - 1) = Worksheets("Loads_Cases").Range("E" & CStr(i + iStart + 1)).Value
'        MyMyName(i - 1) = Worksheets("Loads_Cases").Range("F" & CStr(i + iStart + 1)).Value
'        MySF(i - 1) = Worksheets("Loads_Cases").Range("G" & CStr(i + iStart + 1)).Value
'    Next i
'
'    ' ret = mySapModel.LoadCases.StaticNonlinearStaged ('name, 'number of stages,      'number of operations
'    ret = mySapModel.LoadCases.StaticNonlinearStaged.SetStageData_1("DEAD+OTHERS+PROT", 1, nOperations, MyOperation, MyObjectType, MyObjectName, MyAge, MyMyType, MyMyName, MySF)
    
    '*********************************************************************************************
    ' LOAD CASE STAGE DATA 2 (DEAD + OTHERS (PAV, GUARDA RODAS, DUTOS) + PROTENSAO) + ENC ********
        
    iStart = 1 ' indicates the start of DEAD+OTHERS+PROT+ENC load case (of the header)
    nOperations = Worksheets("Loads_Cases").Range("K" & CStr(iStart)).Value 'operations inside the stage definition, such as Add Structure, Load Objects,...
    
    ReDim MyOperation(nOperations - 1)
    ReDim MyObjectType(nOperations - 1)
    ReDim MyObjectName(nOperations - 1)
    ReDim MyAge(nOperations - 1)
    ReDim MyMyType(nOperations - 1)
    ReDim MyMyName(nOperations - 1)
    ReDim MySF(nOperations - 1)
        
    For i = 1 To nOperations
        MyOperation(i - 1) = Worksheets("Loads_Cases").Range("A" & CStr(i + iStart + 1)).Value
        MyObjectType(i - 1) = Worksheets("Loads_Cases").Range("B" & CStr(i + iStart + 1)).Value
        MyObjectName(i - 1) = Worksheets("Loads_Cases").Range("C" & CStr(i + iStart + 1)).Value
        MyAge(i - 1) = Worksheets("Loads_Cases").Range("D" & CStr(i + iStart + 1)).Value
        MyMyType(i - 1) = Worksheets("Loads_Cases").Range("E" & CStr(i + iStart + 1)).Value
        MyMyName(i - 1) = Worksheets("Loads_Cases").Range("F" & CStr(i + iStart + 1)).Value
        MySF(i - 1) = Worksheets("Loads_Cases").Range("G" & CStr(i + iStart + 1)).Value
    Next i
      
    ' ret = mySapModel.LoadCases.StaticNonlinearStaged ('name, 'number of stages,      'number of operations
    ret = mySapModel.LoadCases.StaticNonlinearStaged.SetStageData_1("DEAD+OTHERS+PROT+ENC", 1, nOperations, MyOperation, MyObjectType, MyObjectName, MyAge, MyMyType, MyMyName, MySF)
    
    
    '*********************************************************************************************
 '   ' LOAD CASE STAGE DATA 3 (DEAD + PAV + PROTENSAO) ********
 '
 '   iStart = 23 ' indicates the start of DEAD+OTHERS+PROT load case (number of the row of the header)
 '   nOperations = Worksheets("Loads_Cases").Range("K" & CStr(iStart)).Value 'operations inside the stage definition, such as Add Structure, Load Objects,...
 '
 '   ReDim MyOperation(nOperations - 1)
 '   ReDim MyObjectType(nOperations - 1)
 '   ReDim MyObjectName(nOperations - 1)
 '   ReDim MyAge(nOperations - 1)
 '   ReDim MyMyType(nOperations - 1)
 '   ReDim MyMyName(nOperations - 1)
 '   ReDim MySF(nOperations - 1)
 '
 '   For i = 1 To nOperations
 '       MyOperation(i - 1) = Worksheets("Loads_Cases").Range("A" & CStr(i + iStart + 1)).Value
 '       MyObjectType(i - 1) = Worksheets("Loads_Cases").Range("B" & CStr(i + iStart + 1)).Value
 '       MyObjectName(i - 1) = Worksheets("Loads_Cases").Range("C" & CStr(i + iStart + 1)).Value
 '       MyAge(i - 1) = Worksheets("Loads_Cases").Range("D" & CStr(i + iStart + 1)).Value
 '       MyMyType(i - 1) = Worksheets("Loads_Cases").Range("E" & CStr(i + iStart + 1)).Value
 '       MyMyName(i - 1) = Worksheets("Loads_Cases").Range("F" & CStr(i + iStart + 1)).Value
 '       MySF(i - 1) = Worksheets("Loads_Cases").Range("G" & CStr(i + iStart + 1)).Value
 '   Next i
 '
 '   ' ret = mySapModel.LoadCases.StaticNonlinearStaged ('name, 'number of stages,      'number of operations
 '   ret = mySapModel.LoadCases.StaticNonlinearStaged.SetStageData_1("DEAD+PAV+PROT", 1, nOperations, MyOperation, MyObjectType, MyObjectName, MyAge, MyMyType, MyMyName, MySF)
     
    '*********************************************************************************************
    ' ADD CABLE LOADs
    ' CABLE TARGET LOADS
    
    Dim node As String 'name of the cable
    Dim force As Double 'force that is applied in the node to emulate the cable
    Dim PointLoadValue() As Double 'force vector used in SAP2000 (fx, fy, fz, Mx, My, Mz)
    ReDim PointLoadValue(5)
    
    For i = 1 To nCables 'assign cable target force
        force = Worksheets("Stays").Range("O" & CStr(i + 1)).Value
        node = Worksheets("Stays").Range("M" & CStr(i + 1)).Value
        ret = mySapModel.CableObj.SetLoadTargetForce(node, "PROT", force, 0)
    Next i

    '*********************************************************************************************
    ' ADD LOAD PAVIMENTACAO, GUARDA RODAS, DUTO PLACAS E TUBOS
    
    Dim pavLoad As Double
    Dim rhoPav As Double
    Dim asphaltThickness As Double
    Dim LanesWidth As Double
    
    'add pavimentacao load
    asphaltThickness = Worksheets("Loads_Patterns").Range("C5").Value ' 0.115  ' m
    rhoPav = Worksheets("Loads_Patterns").Range("C6").Value '23.5 ' asphalt density kN/m3
    LanesWidth = Worksheets("Loads_Patterns").Range("C7").Value ' 10.5 ' m
    pavLoad = rhoPav * asphaltThickness * LanesWidth ' 23.53596 kN/m3 * 0.2 m * 10.5 m = pavimentacao kN/m
    
    'add guarda Rodas and Ducts load
    guardaRodasLoad = Worksheets("Loads_Patterns").Range("C11").Value  ' ~5.6 kN/m -> per Guarda Rodas
    dutoPlacaTubLoad = Worksheets("Loads_Patterns").Range("C14").Value ' kN/m -> ducts, plates and tubes
    
    For i = 1 To nFrames                                 'Name,'load pattern, 1=kN/m, 10=Gravity direction, distance to start, dist to from start to end, start load value kN/m, end load value kN/m
        ret = mySapModel.FrameObj.SetLoadDistributed(FrameName(i - 1), "PAVIMENTACAO", 1, 10, 0, 1, pavLoad, pavLoad) ' add pavimentacao to all frames
        ret = mySapModel.FrameObj.SetLoadDistributed(FrameName(i - 1), "GUARDARODAS", 1, 10, 0, 1, guardaRodasLoad, guardaRodasLoad)  ' add guarda rodas load to all frames
        ret = mySapModel.FrameObj.SetLoadDistributed(FrameName(i - 1), "DUTOPLACATUBOS", 1, 10, 0, 1, dutoPlacaTubLoad, dutoPlacaTubLoad)  ' add plates, tubes, ducts to all frames
        'ret = mySapModel.FrameObj.SetLoadDistributed(FrameName(1)        , "4"       , 1 , 11, 0, 1, 2, 2)
    Next i


    '*********************************************************************************************
    ' ADD LOAD ENCHIMENTOS
    
    Dim frame As Integer ' frame name in which the load is applied
    Dim loadEnc_i As Double ' initial load applied in the frame
    Dim loadEnc_f As Double ' final load applied in the frame
    Dim nEnc As Integer ' number of Enchimentos
    
    nEnc = Worksheets("Loads_Patterns").Range("H18").Value ' number of Enchimentos
    
    For i = 1 To nEnc                                 'Name,'load pattern, 1=kN/m, 10=Gravity direction, distance to start, dist to end,load value kN/m, value M/m
        frame = Worksheets("Loads_Patterns").Range("A" & CStr(i + 17)).Value
        loadEnc_i = Worksheets("Loads_Patterns").Range("D" & CStr(i + 17)).Value
        loadEnc_f = Worksheets("Loads_Patterns").Range("E" & CStr(i + 17)).Value
        ret = mySapModel.FrameObj.SetLoadDistributed(frame, "ENCHIMENTOS", 1, 10, 0, 1, loadEnc_i, loadEnc_f)   ' add the load Enchimento in the defined frames in the Excel Spreadsheet
    Next i
 

    'assign loading for load pattern 4
    'ret = mySapModel.FrameObj.SetLoadDistributed(FrameName(1), "4", 1, 11, 0, 1, 2, 2)
    'assign loading for load pattern 5
    'ret = mySapModel.FrameObj.SetLoadDistributed(FrameName(0), "5", 1, 2, 0, 1, 2, 2, "Local")
    'ret = mySapModel.FrameObj.SetLoadDistributed(FrameName(1), "5", 1, 2, 0, 1, -2, -2, "Local")
    'assign loading for load pattern 6
    'ret = mySapModel.FrameObj.SetLoadDistributed(FrameName(0), "6", 1, 2, 0, 1, 0.9984, 0.3744, "Local")
    'ret = mySapModel.FrameObj.SetLoadDistributed(FrameName(1), "6", 1, 2, 0, 1, -0.3744, 0, "Local")
    'assign loading for load pattern 7
    'ret = mySapModel.FrameObj.SetLoadPoint(FrameName(1), "7", 1, 2, 0.5, -15, "Local")

    '*********************************************************************************************
    ' SAVE MODEL
    ' The ModelPath is defined at the beggining
    
    ret = mySapModel.File.Save(ModelPath)

  
    '*********************************************************************************************
    ' RUN ANALYSIS
    ' This will create the analysis model

    ret = mySapModel.Analyze.RunAnalysis

    
    '*********************************************************************************************
    ' GET RESULTS
    ' initialize for results
    ' then gets the results and write them in the Excel spreadsheet
    
    Dim nNodes As Integer
    Dim NumberResults As Long
    Dim Obj() As String
    Dim Elm() As String
    Dim LoadCase() As String
    Dim StepType() As String
    Dim StepNum() As Double
    Dim U1() As Double
    Dim U2() As Double
    Dim U3() As Double
    Dim R1() As Double
    Dim R2() As Double
    Dim R3() As Double
    
    nNodes = nFrames + 1 'frame deck nodes

'    ' Get results of the load case 1 (DEAD + OTHERS (PAV, GUARDA RODAS, DUTOS) + PROTENSAO)
'    ret = mySapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
'    ret = mySapModel.Results.Setup.SetCaseSelectedForOutput("DEAD+OTHERS+PROT")
'    ret = mySapModel.Results.JointDispl("nodes_DECK", eItemTypeElm_GroupElm, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, U1, U2, U3, R1, R2, R3)
'    Worksheets("Results_Case1").Range("D2:D" & CStr(2 * (nNodes + 1))).Value = 0 ' delete old results
'    Worksheets("Results_Case1").Range("D2:D" & CStr(2 * (nNodes + 1))).Value = Application.Transpose(U3)
    
    ' Get results of the load case 2 (DEAD + OTHERS (PAV, GUARDA RODAS, DUTOS) + PROTENSAO + ENCHIMENTOS)
    ret = mySapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
    ret = mySapModel.Results.Setup.SetCaseSelectedForOutput("DEAD+OTHERS+PROT+ENC")
    ret = mySapModel.Results.JointDispl("nodes_DECK", eItemTypeElm_GroupElm, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, U1, U2, U3, R1, R2, R3)
    Worksheets("Results_Case2").Range("D2:D" & CStr(2 * (nNodes + 1))).Value = 0 ' delete old results
    Worksheets("Results_Case2").Range("D2:D" & CStr(2 * (nNodes + 1))).Value = Application.Transpose(U3)
    
 '   ' Get results of the load case 3 (DEAD + PAV + PROTENSAO)
 '   ret = mySapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
 '   ret = mySapModel.Results.Setup.SetCaseSelectedForOutput("DEAD+PAV+PROT")
 '   ret = mySapModel.Results.JointDispl("nodes_DECK", eItemTypeElm_GroupElm, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, U1, U2, U3, R1, R2, R3)
 '   Worksheets("Results_Case3").Range("D2:D" & CStr(2 * (nNodes + 1))).Value = 0 ' delete old results
 '   Worksheets("Results_Case3").Range("D2:D" & CStr(2 * (nNodes + 1))).Value = Application.Transpose(U3)


    '*********************************************************************************************
    ' LOOP through all data entry (to do a Monte Carlo Simulation)
    ' j is the number of the cell that will be read and also where it will be saved
    Dim startTime As Double
    Dim secondsElapsed As String
    
    startTime = Timer
    
    For j = 1 To 100
        'Unlock model
        ret = mySapModel.SetModelIsLocked(False) 'unlock the model
        
        'delete cable target force
        ret = mySapModel.CableObj.DeleteLoadTargetForce("ALL", "PROT")
        
        ' ADD CABLE LOADs
        ' CABLE TARGET LOADS
        For i = 1 To nCables 'assign cable target force
            force = Worksheets("Sheet1").Range("A1").Offset(j - 1, i - 1).Value
            node = Worksheets("Stays").Range("M" & CStr(i + 1)).Value
            ret = mySapModel.CableObj.SetLoadTargetForce(node, "PROT", force, 0)
        Next i

        ret = mySapModel.File.Save(ModelPath) ' SAVE MODEL
        ret = mySapModel.Analyze.RunAnalysis ' run analysis in the loop
        ret = mySapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
        ret = mySapModel.Results.Setup.SetCaseSelectedForOutput("DEAD+OTHERS+PROT+ENC")
        
        ret = mySapModel.Results.JointDispl("nodes_DECK", eItemTypeElm_GroupElm, NumberResults, Obj, Elm, LoadCase, StepType, StepNum, U1, U2, U3, R1, R2, R3)
        
        Worksheets("Results").Range("A" & CStr(j) & ":CL" & CStr(j)).Value = U3
        
        Debug.Print j
        
    Next j

    secondsElapsed = Round(Timer - startTime, 2)
    MsgBox "The time to run one analysis is " & secondsElapsed & " seconds", vbInformation
    
    '*********************************************************************************************
    ' CLOSE APPLICATION
    
    'mySapObject.ApplicationExit False
    Set mySapModel = Nothing
    Set mySapObject = Nothing

End Sub


