// ReSharper disable InconsistentNaming
namespace Automation.GenerateIgxl.HardIp.InputObject
{
    public enum RegisterAssignType
    {
        //SweepTable, 
        NoneType,
        CUS_Str_DigCapData, CUS_Str_MainProgram,
        DigSrc_Equation, DigSrc_Assignment, DigSrc_FlowForLoopIntegerName,
        Calc_Eqn, Interpose_PreMeas,
        MeasV_PinS, Meas_StoreName, MeasI_PinS, MeasF_PinS_SingleEnd, MeasI_Range, TrimDictionaryStoreName,
        MeasPins, MeasName,
        ForceV_Val, ForceI_Val, TestSequence,
        CalcEquName, CalcStoreName, //LCD
        HIP_eFuse_Read, HIP_eFuse_Write  //Efuse Read/Write
    }
}
