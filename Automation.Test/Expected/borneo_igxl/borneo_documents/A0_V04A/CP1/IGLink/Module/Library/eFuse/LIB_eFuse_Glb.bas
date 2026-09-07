Attribute VB_Name = "LIB_eFuse_Glb"
#Const isUFP = True
Option Explicit
'' special for sepvm---------------------------------------
''===sepvm===
Public Type SEPVM_CONST_Type
    Lo_Limit As Double
    Hi_Limit As Double
    FuseValue0 As Variant
End Type

Public Type SEPVM_SHEET_Type
    FuseName As String
    Sepvm_Const() As SEPVM_CONST_Type
    RangeCnt As Long
End Type

'20220614, Add for parsing sram_soc table dynamic
Public Sepvm_table_Arr() As SEPVM_SHEET_Type

''===shmoo===
Public gS_LVCC_FuseValue As New SiteVariant

''BinCut printout??
Public Const gB_newDlog_Flag = True

''for removing orignal efuse------------------------------------------------
'BKM variable
Public gS_BKM_Unknown As String
Public gS_BKM_IEDA As String
Public gS_BKMPreName As String '' Save current lot ID, If Lot ID is the same then no need to read BKM again

''overflow variable
Public gB_Fuse_Skip As New SiteBoolean

''disable efuse limit check
Public gB_eFuse_Disable_ChkLMT_Flag As Boolean

''Add for disable printing Efuse Detail Datalog
Public gB_eFuse_Disable_DSPwavePrint_Flag As Boolean
Public gB_eFuse_Disable_DecodeDataPrint_Flag As Boolean
Public ForceDecodeEnable As Boolean

''20220524, modify for syntax check only check current job fields
Public gB_eFuse_Disable_SyntaxCheckAll_Flag As Boolean

'20220106, Add for Real value validation( reverse bit)
Public gB_efuse_DicValue_Chk_Flag As Boolean
''for removing orignal efuse------------------------------------------------end

'20220803, Add for single double bit debug print
Public gB_efuse_DebugPrint_SingleDoubleBits_Flag As Boolean
