namespace RF_PatternTool.PatternStruct
{
    public class SubRoutine
    {
        public List<string> RawDataList = new List<string>();
        public SubRoutine() { }
        //Read Raw data template
        public void ReadTemplate()
        {
            List<string> lines = Template.TemplateSet.SubrTemp.Replace("\r", "").Split('\n').ToList();
            foreach (string line in lines)
            {
                if (line.Contains(">"))
                {
                    RawDataList.Add(line);
                }
                else
                {
                    RawDataList.Add(line);
                }
            }
        }
        //srm_vector => PP_BTCA0_S_PLLP_AN_RFD1_PFF_JTG_UNS_ALLFRV_SI_XTALLDO1_VOLTAGE_srm_meas
        //global subr PP_BTCA0_S_PLLP_AN_RFD1_PFF_JTG_UNS_ALLFRV_SI_XTALLDO1_VOLTAGE_digsrc
        public void Write(List<string> patrows, string pattern, string subrName)
        {
            if (patrows.Count == 0)
            {
                patrows.Add("opcode_mode = single;");
                patrows.Add("import tset tsetJTAG;");
                patrows.Add("instruments = {");
                patrows.Add("(JTAG_TDO):DigCap 1:lsb:serial:auto_trig_enable;");
                patrows.Add("(JTAG_TDI):DigSrc 1:lsb:serial;");
                patrows.Add("}");
            }
            foreach (string data in RawDataList)
            {
                if (data.Contains("name"))
                {
                    string newName = pattern;
                    if (data.Contains("srm_vector"))
                    {
                        ;
                    }
                    else if (data.Contains("global"))
                    {
                        newName = subrName;
                    }

                    patrows.Add(data.Replace("name", newName));
                }
                else
                {
                    patrows.Add(data);
                }
            }
        }


    }
}
