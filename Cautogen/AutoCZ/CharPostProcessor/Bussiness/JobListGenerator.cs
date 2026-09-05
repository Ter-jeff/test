using System.IO;
using System.Text;

namespace Cautogen.AutoCZ.CharPostProcessor.Bussiness
{
    public class JobListGenerator
    {
        private readonly StreamWriter _igxlWriter;

        public void WriteHeader()
        {
            const string header = "DTJobListSheet,version=2.5:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1	Job List";
            _igxlWriter.WriteLine(header);
            _igxlWriter.WriteLine();
            WriteColumnsHeader();
        }

        private void WriteColumnsHeader()
        {
            var firstRow = new StringBuilder();
            var secondRow = new StringBuilder();
            firstRow.Append("\t\tSheet Parameters\t\t");
            secondRow.Append("\tJob Name\tPin Map\tTest Instances\tFlow Table\tAC Specs\tDC Specs\tPattern Sets\tPattern Groups\tBin Table\tCharacterization\tTest Procedures\tMixed Signal Timing\tWave Definitions\tPsets\tSignals\tPort Map\tFractional Bus\tConcurrent Sequence\tComment\t");
            _igxlWriter.WriteLine(firstRow.ToString());
            _igxlWriter.WriteLine(secondRow.ToString());
        }

        public void WriteRows(string jobName, string pinMap, string testInstance, string flowTable, string acSpecs, string dcSpecs, string patternSets,
            string patternGroups, string binTable, string characterization, string testProcedures, string mixedSignalTiming, string waveDefinition,
            string psets, string signals, string portMap, string fractionalBus, string concurrentSequence, string comment)
        {
            var jobRow = new StringBuilder();
            jobRow.Append("\t");
            jobRow.Append(jobName);
            jobRow.Append("\t");
            jobRow.Append(pinMap);
            jobRow.Append("\t");
            jobRow.Append(testInstance);
            jobRow.Append("\t");
            jobRow.Append(flowTable);
            jobRow.Append("\t");
            jobRow.Append(acSpecs);
            jobRow.Append("\t");
            jobRow.Append(dcSpecs);
            jobRow.Append("\t");
            jobRow.Append(patternSets);
            jobRow.Append("\t");
            jobRow.Append(patternGroups);
            jobRow.Append("\t");
            jobRow.Append(binTable);
            jobRow.Append("\t");
            jobRow.Append(characterization);
            jobRow.Append("\t");
            jobRow.Append(testProcedures);
            jobRow.Append("\t");
            jobRow.Append(mixedSignalTiming);
            jobRow.Append("\t");
            jobRow.Append(waveDefinition);
            jobRow.Append("\t");
            jobRow.Append(psets);
            jobRow.Append("\t");
            jobRow.Append(signals);
            jobRow.Append("\t");
            jobRow.Append(portMap);
            jobRow.Append("\t");
            jobRow.Append(fractionalBus);
            jobRow.Append("\t");
            jobRow.Append(concurrentSequence);
            jobRow.Append("\t");
            jobRow.Append(comment);
            _igxlWriter.WriteLine(jobRow.ToString());
        }

        public JobListGenerator(string fileName)
        {
            _igxlWriter = new StreamWriter(fileName, false);  // overwrite the jobSheet if already exists
        }

        public void Close()
        {
            _igxlWriter.Close();
        }
    }
}
