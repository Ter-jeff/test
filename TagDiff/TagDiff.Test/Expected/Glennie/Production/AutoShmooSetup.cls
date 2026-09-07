VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoShmooSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum ShmooPinTypes
    DC = 1
    AC = 2
    Nwrire = 3
    None = 0
End Enum

Private Enum ShmooLimits
    highLimit = 1
    LowLimit = 2
    None = 0
End Enum

'========== const =========='
Private Const Char_Setup_Name = "AutoShmoo" ' shmoo setup name
Private Const Core_power = "CorePower" ' shmoo core power
Private Const SHMOO_INFO = "PrintShmooInfo" ' shmoo info to print shmoo row data
Private Const ENHANCE = "Enhanced" ' shmoo output format

Private Const SHIFT_IN = "Shift_In" ' shmoo axis name
Private Const NWire = "XI0" ' shmoo axis name
Private Const FREERUN_START = "freerunclk_set_XY" ' nwire interpose start
Private Const FREERUN_STOP = "freerunclk_stop" ' nwire interpose stop
Private Const FREERUN_PORT = "XI0_Diff_Port" ' nwire port

Private Const AC_TYPE = "AC Spec"
Private Const GLOBAL_TYPE = "Global Spec"
Private Const GLOBAL_XPIN = "Shmoo_Glb_GLB"
Private Const GLOBAL_YPIN = "Shmoo_GLB"

Private Const ONE_DIMENSION_ONLY = False ' shmoo 1D only

Private Const SAFE_VOLTAGE = 0.775 ' safe voltage
Private Const DEFAULT_SHIFT = 50000000 ' default shift frequency (50M)
Private Const DEFAULT_NWIRE = 24000000 ' defautl nwire frequency (24M)
Private Const SHMOO_VOLTAGE_BUFFER = 0.3 ' for voltage shmoo range buffer
Private Const SHMOO_SHIFT_BUFFER = 40000000 ' for shift shmoo range buffer (40M)
Private Const SHMOO_NWIRE_BUFFER = 4000000 ' for nwire shmoo range buffer (4M)
Private Const SHMOO_VOLTAGE_HIGH_LIMIT = 1.2 ' shmoo voltage sweep high limit
Private Const SHMOO_VOLTAGE_LOW_LIMIT = 0.4 ' shmoo voltage sweep low limit
Private Const SHMOO_SHIFT_HIGH_LIMIT = 250000000 ' shmoo shift sweep high limit (250M)
Private Const SHMOO_SHIFT_LOW_LIMIT = 10000000 ' shmoo shift sweep low limit (10M)
Private Const SHMOO_NWIRE_HIGH_LIMIT = 28000000 ' shmoo nwire sweep high limit (28M)
Private Const SHMOO_NWIRE_LOW_LIMIT = 20000000 ' shmoo nwire sweep low limit (20M)
Private Const SHMOO_VOLTAGE_STEP = 0.02 ' shmoo voltage sweep step
Private Const SHMOO_SHIFT_STEP = 10000000 ' shmoo shift sweep step (10M)
Private Const SHMOO_NWIRE_STEP = 1000000 ' shmoo nwire sweep step (1M)
Private Const SHMOO_JUMP_STEP = 6 ' shmoo jump step

Private Const IS_LINEAR_SEARCH = False ' default set to jump

'========== fields =========='
Private charSetupName_ As String
Private shmooTestPattern_ As String
Private glbShmooTestPattern_ As String


'========== constructor =========='
Sub Class_Initialize()
    'CreateShmooSetup
End Sub

'========== properties =========='
Public Property Get CharSetupName() As String
    CharSetupName = charSetupName_
End Property
Public Property Get IsValid() As Boolean
    IsValid = TheExec.DevChar.Setups(charSetupName_).IsValid
End Property
Public Property Get GetXaxisGlobalName() As String
    GetXaxisGlobalName = GLOBAL_XPIN
End Property
Public Property Get GetYaxisGlobalName() As String
    GetYaxisGlobalName = GLOBAL_YPIN
End Property
Public Property Get IsOneDimensionOnly() As Boolean
    IsOneDimensionOnly = ONE_DIMENSION_ONLY
End Property
Public Property Let ShmooTestPattern(patternSet As String)
    shmooTestPattern_ = patternSet
End Property

'========== methods =========='
Private Function CreateShmooSetup()
    ''' create template shmoo for default setting '''
    
    ''' check if setup already exist, if so, reset it
    If TheExec.DevChar.Setups.Contains(charSetupName_) Then
        TheExec.DevChar.Setups.Remove charSetupName_
    End If
    TheExec.DevChar.Setups.Add charSetupName_
    
    ''' setup default setting
    With TheExec.DevChar.Setups(charSetupName_)
        .TestMethod = tlDevCharTestMethod_Retest
        .Shmoo.PinsExecMode = tlDevCharPinExecMode_Simultaneous
        .Shmoo.InterposeFunctions.PostStep.Arguments = Core_power
        .Shmoo.InterposeFunctions.PostStep.name = SHMOO_INFO
        .Output.Format = ENHANCE
        .Output.Destinations.Datalog = tlDevCharOutputDestinationState_Enabled
    End With
    
End Function
Private Function GetAlgorithmType(axisType As tlDevCharShmooAxis) As tlDevCharShmooPGA
    ''' get shmoo algorithm type '''
    
    Select Case axisType
        Case tlDevCharShmooAxis_X:
            If IS_LINEAR_SEARCH Then
                GetAlgorithmType = tlDevCharShmooPGA_Linear
            Else
                GetAlgorithmType = tlDevCharShmooPGA_Jump
            End If
        Case Else:
            GetAlgorithmType = tlDevCharShmooPGA_Linear
    End Select
End Function
Private Function AxisTypeName(axisType As tlDevCharShmooAxis) As String
    ''' get corresponding shmoo axis name for nwire interpose '''
    
    Select Case axisType
        Case tlDevCharShmooAxis_X:
            AxisTypeName = "X"
        Case tlDevCharShmooAxis_Y:
            AxisTypeName = "Y"
        Case tlDevCharShmooAxis_Z:
            AxisTypeName = "Z"
    End Select
End Function
Private Function GetShmooGlobalName(axisType As tlDevCharShmooAxis) As String
    ''' get shmoo global name for voltage sweep '''
    
    Select Case axisType
        Case tlDevCharShmooAxis_X:
            GetShmooGlobalName = GLOBAL_XPIN
        Case tlDevCharShmooAxis_Y:
            GetShmooGlobalName = GLOBAL_YPIN
        Case tlDevCharShmooAxis_Z:
            ' should not be the case
    End Select
End Function
Private Function CheckPinType(pin As String) As ShmooPinTypes
    ''' check the shmoo pin type for shmoo axis creation '''
    
    If pin Like "VDD*" Then
        CheckPinType = ShmooPinTypes.DC
    ElseIf pin Like "*XI0*" Then
        CheckPinType = ShmooPinTypes.Nwrire
    Else
        CheckPinType = ShmooPinTypes.AC
    End If
End Function
Private Function GetMaxMinValue(data As SiteDouble, limit As ShmooLimits) As Double
    ''' get maximum/minmum value from site double '''
    Dim site As Variant
    
    For Each site In TheExec.sites.Selected
        Select Case limit
            Case ShmooLimits.highLimit:
                GetMaxMinValue = Max(GetMaxMinValue, Format(data(site), "0.00"))
            Case ShmooLimits.LowLimit:
                GetMaxMinValue = Min(Max(GetMaxMinValue, 1E+308), Format(data(site), "0.00"))
        End Select
    Next site
End Function
Private Function CalculateShmooLimit(Condition As Double, limit As ShmooLimits, ShmooPinType As ShmooPinTypes) As Double
    ''' calculate high/low limit for shmoo range. Force to safe voltage if limit not found '''
    
    Select Case ShmooPinType
        Case ShmooPinTypes.DC
            Select Case limit
                Case ShmooLimits.highLimit:
                    CalculateShmooLimit = IIf(Condition + SHMOO_VOLTAGE_BUFFER >= SHMOO_VOLTAGE_HIGH_LIMIT, SHMOO_VOLTAGE_HIGH_LIMIT, Condition + SHMOO_VOLTAGE_BUFFER)
                Case ShmooLimits.LowLimit:
                    CalculateShmooLimit = IIf(Condition - SHMOO_VOLTAGE_BUFFER <= SHMOO_VOLTAGE_LOW_LIMIT, SHMOO_VOLTAGE_LOW_LIMIT, Condition - SHMOO_VOLTAGE_BUFFER)
                Case Else:
                    ' something wrong here
                    CalculateShmooLimit = SAFE_VOLTAGE
            End Select
        Case ShmooPinTypes.AC:
            Select Case limit
                Case ShmooLimits.highLimit:
                    CalculateShmooLimit = IIf(Condition + SHMOO_SHIFT_BUFFER >= SHMOO_SHIFT_HIGH_LIMIT, SHMOO_SHIFT_HIGH_LIMIT, Condition + SHMOO_SHIFT_BUFFER)
                Case ShmooLimits.LowLimit:
                    CalculateShmooLimit = IIf(Condition - SHMOO_SHIFT_BUFFER <= SHMOO_SHIFT_LOW_LIMIT, SHMOO_SHIFT_LOW_LIMIT, Condition - SHMOO_SHIFT_BUFFER)
                Case Else:
                    ' something wrong here
                    CalculateShmooLimit = DEFAULT_SHIFT
            End Select
        Case ShmooPinTypes.Nwrire:
            Select Case limit
                Case ShmooLimits.highLimit:
                    CalculateShmooLimit = IIf(Condition + SHMOO_NWIRE_BUFFER >= SHMOO_NWIRE_HIGH_LIMIT, SHMOO_NWIRE_HIGH_LIMIT, Condition + SHMOO_NWIRE_BUFFER)
                Case ShmooLimits.LowLimit:
                    CalculateShmooLimit = IIf(Condition - SHMOO_NWIRE_BUFFER <= SHMOO_NWIRE_LOW_LIMIT, SHMOO_NWIRE_LOW_LIMIT, Condition - SHMOO_NWIRE_BUFFER)
                Case Else:
                    ' something wrong here
                    CalculateShmooLimit = DEFAULT_NWIRE
            End Select
    End Select
End Function
Private Function AddAxis(pin As String, Condition As SiteDouble, axisType As tlDevCharShmooAxis)
    ''' add shmoo axis into characterization '''
    With TheExec.DevChar.Setups.item(charSetupName_).Shmoo.axes
        Select Case CheckPinType(pin)
            Case ShmooPinTypes.DC:
                ''' add pin
                .Add axisType, Replace(pin, ",", "")
                
                ''' shmoo pin setup
                .item(axisType).parameter.type = GLOBAL_TYPE
                .item(axisType).parameter.name = GetShmooGlobalName(axisType)
                .item(axisType).ApplyTo.pins = pin
                .item(axisType).parameter.range.StepSize = SHMOO_VOLTAGE_STEP
                
                ' shmoo range setting (from high to low to prevent tip burnt)
                .item(axisType).parameter.range.from = CalculateShmooLimit(GetMaxMinValue(Condition, highLimit), highLimit, ShmooPinTypes.DC)
                .item(axisType).parameter.range.to = CalculateShmooLimit(GetMaxMinValue(Condition, LowLimit), LowLimit, ShmooPinTypes.DC)
            Case ShmooPinTypes.AC:
                ''' add pin
                .Add axisType, SHIFT_IN
                
                ''' shmoo pin setup
                .item(axisType).parameter.type = AC_TYPE
                .item(axisType).parameter.name = pin
                .item(axisType).parameter.range.StepSize = SHMOO_SHIFT_STEP
                
                ' shmoo range setting (from low to high)
                .item(axisType).parameter.range.from = CalculateShmooLimit(GetMaxMinValue(Condition, LowLimit), LowLimit, ShmooPinTypes.AC)
                .item(axisType).parameter.range.to = CalculateShmooLimit(GetMaxMinValue(Condition, highLimit), highLimit, ShmooPinTypes.AC)
            Case ShmooPinTypes.Nwrire:
                ''' add pin
                .Add axisType, NWire
                
                ''' shmoo pin setup
                .item(axisType).parameter.type = AC_TYPE
                .item(axisType).parameter.name = pin
                .item(axisType).parameter.range.StepSize = SHMOO_NWIRE_STEP
                
                ''' special setup for nwrie shmoo
                .item(axisType).InterposeFunctions.PrePoint = FREERUN_START
                .item(axisType).InterposeFunctions.PostPoint = FREERUN_STOP
                .item(axisType).InterposeFunctions.PrePoint.Arguments = AxisTypeName(axisType) & "," & FREERUN_PORT & "," & pin
                .item(axisType).InterposeFunctions.PostPoint.Arguments = FREERUN_PORT
                
                ' shmoo range setting (from low to high)
                .item(axisType).parameter.range.from = CalculateShmooLimit(GetMaxMinValue(Condition, LowLimit), LowLimit, ShmooPinTypes.Nwrire)
                .item(axisType).parameter.range.to = CalculateShmooLimit(GetMaxMinValue(Condition, highLimit), highLimit, ShmooPinTypes.Nwrire)
            Case Else:
                ' something wrong here
        End Select
        
        ' shmoo search setup. default jump with 6 step
        .item(axisType).Algorithm.type = GetAlgorithmType(axisType)
        .item(axisType).Algorithm.Arguments = SHMOO_JUMP_STEP
        .item(axisType).parameter.range.CalculatedField = tlDevCharRangeField_Steps
        
    End With
End Function
Public Function SetupShmoo(ShmooItem As AutoShmooItem, charName As String)
    ''' setup characterization for test item '''
    charSetupName_ = Char_Setup_Name + "_Site" + charName
    CreateShmooSetup
    
    ''' 1-D setup
    AddAxis ShmooItem.powerPin, ShmooItem.OperationPoint, tlDevCharShmooAxis_X
    'TODO need to consider track pin?

    ''' 2-D setup
    If Not ONE_DIMENSION_ONLY Then
        AddAxis ShmooItem.SecondPin, ShmooItem.SecondOperationPoint, tlDevCharShmooAxis_Y
    End If
End Function

Public Function SetShmooTestPattern()
    glbShmooTestPattern_ = Shmoo_Pattern
    Shmoo_Pattern = shmooTestPattern_
End Function

Public Function RestoreShmooTestPattern()
    Shmoo_Pattern = glbShmooTestPattern_
End Function

