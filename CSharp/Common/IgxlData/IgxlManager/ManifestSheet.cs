﻿using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IgxlData.IgxlManager
{
    public class ManifestSheet
    {
        public string GeneratedType;
        public List<ManifestSheetItem> Items = new List<ManifestSheetItem>();
        public string JobName;
        public string MainFlow = "Not Specified";
        public string ProjectName = "";
        public string RelativeType;

        public Dictionary<string, List<ManifestSheetItem>> SubPrograms =
            new Dictionary<string, List<ManifestSheetItem>>();

        public string TargetFolder;

        public ManifestSheet(string targetFolder)
        {
            TargetFolder = targetFolder;
        }


        public void UpdateManifestHeaderInfo(string fileInfo)
        {
            ProjectName = Path.GetFileNameWithoutExtension(fileInfo.Split(',')[0]);
            GeneratedType = fileInfo.Split(',')[1].Split('.').ToList().Last();
            JobName = fileInfo.Split(',')[2];
        }

        public bool IsValidManifestItem(List<string> header)
        {
            if (header[0] != "Sub Program") return false;
            if (header[1] != "Source File") return false;
            if (header[2] != "Imported Name") return false;
            return true;
        }

        public void UpdateManifestItem(List<string> content, string iglink)
        {
            var item = new ManifestSheetItem();
            item.Subprogram = content[0];
            item.FileName = content[2];
            item.FullFilePath = Path.Combine(iglink, content[1]);
            item.AbsoluteFilePath = Path.Combine(TargetFolder, item.FullFilePath);
            item.RelatedPath = content[1];
            Items.Add(item);
        }

        public void UpdateManifestItemWithUnknownItem(string subprogram, string filename, string relatedPath,
            string fullPath)
        {
            var item = new ManifestSheetItem();
            item.Subprogram = subprogram;
            item.FileName = filename;
            item.FullFilePath = fullPath;
            item.AbsoluteFilePath = Path.Combine(TargetFolder, item.FullFilePath);
            item.RelatedPath = relatedPath;
            Items.Add(item);
        }

        public void ArrangeSubprogramItems()
        {
            var groups = Items.GroupBy(p => p.Subprogram).ToDictionary(p => p.Key, p => p.ToList());
            //Remove "Generated by IG-Link" items
            foreach (var group in groups)
                foreach (var subgroup in group.Key.Split(',').ToList())
                    //if not contains, create new one
                    if (!SubPrograms.ContainsKey(subgroup))
                        SubPrograms.Add(subgroup, group.Value);
                    else
                        SubPrograms[subgroup].AddRange(group.Value);
        }
    }

    public class ManifestSheetItem
    {
        public string AbsoluteFilePath;
        public string FileName;
        public string FullFilePath;
        public string RelatedPath;
        public string Subprogram;
    }
}