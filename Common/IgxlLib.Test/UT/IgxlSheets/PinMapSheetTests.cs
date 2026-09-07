using System;
using System.Collections.Generic;
using System.IO;

using IgxlLib.Const;
using IgxlLib.IgxlBase;
using IgxlLib.IgxlConst;
using IgxlLib.IgxlSheets;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using MockLib;

namespace IgxlLib.Test.UT.IgxlSheets
{
    [TestClass]
    public class PinMapSheetTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            MockService.Mock();
        }

        [TestMethod]
        public void PinMapSheet_Constructor_WithSheetName()
        {
            // Arrange
            string sheetName = "PinMap";

            // Act
            var pinMapSheet = new PinMapSheet(sheetName);

            // Assert
            Assert.IsNotNull(pinMapSheet);
            Assert.AreEqual(sheetName, pinMapSheet.Name);
            Assert.AreEqual("DTPinMap", pinMapSheet.SheetType);
            Assert.AreEqual(IgxlSheetNames.PinMap, pinMapSheet.IgxlSheetName);
        }

        [TestMethod]
        public void PinMapSheet_PinList_InitiallyEmpty()
        {
            // Arrange & Act
            var pinMapSheet = new PinMapSheet("PinMap");

            // Assert
            Assert.AreEqual(0, pinMapSheet.PinList.Count);
        }

        [TestMethod]
        public void PinMapSheet_GroupList_InitiallyEmpty()
        {
            // Arrange & Act
            var pinMapSheet = new PinMapSheet("PinMap");

            // Assert
            Assert.AreEqual(0, pinMapSheet.GroupList.Count);
        }

        [TestMethod]
        public void PinMapSheet_AddPin()
        {
            // Arrange
            var pinMapSheet = new PinMapSheet("PinMap");
            var pin = new Pin { PinName = "Pin1" };

            // Act
            pinMapSheet.AddPin(pin);

            // Assert
            Assert.AreEqual(1, pinMapSheet.PinList.Count);
            Assert.AreEqual("Pin1", pinMapSheet.PinList[0].PinName);
        }

        [TestMethod]
        public void PinMapSheet_AddPins()
        {
            // Arrange
            var pinMapSheet = new PinMapSheet("PinMap");
            var pins = new List<Pin>
            {
                new() { PinName= "Pin1"},
                new() { PinName= "Pin2"},
                new() { PinName= "Pin3"}
            };

            // Act
            pinMapSheet.AddPins(pins);

            // Assert
            Assert.AreEqual(3, pinMapSheet.PinList.Count);
        }

        [TestMethod]
        public void PinMapSheet_AddGroup()
        {
            // Arrange
            var pinMapSheet = new PinMapSheet("PinMap");
            var group = new PinGroup("Group1");
            var pin = new Pin { PinName = "Pin1" };
            group.AddPin(pin);

            // Act
            pinMapSheet.AddGroup(group);

            // Assert
            Assert.AreEqual(1, pinMapSheet.GroupList.Count);
            Assert.AreEqual("Group1", pinMapSheet.GroupList[0].PinName);
        }

        [TestMethod]
        public void PinMapSheet_AddGroups()
        {
            // Arrange
            var pinMapSheet = new PinMapSheet("PinMap");
            var group = new PinGroup("Group1");
            var pin = new Pin { PinName = "Pin1" };
            group.AddPin(pin);
            var groups = new List<PinGroup>
            {
                group,
                group,
                group
            };

            // Act
            pinMapSheet.AddGroups(groups);

            // Assert
            Assert.AreEqual(1, pinMapSheet.GroupList.Count);
        }

        [TestMethod]
        public void PinMapSheet_GetPin_ByName()
        {
            // Arrange
            var pinMapSheet = new PinMapSheet("PinMap");
            var pin = new Pin { PinName = "Pin1" };
            pinMapSheet.AddPin(pin);

            // Act
            Pin retrievedPin = pinMapSheet.GetPin("Pin1");

            // Assert
            Assert.IsNotNull(retrievedPin);
            Assert.AreEqual("Pin1", retrievedPin.PinName);
        }

        [TestMethod]
        public void PinMapSheet_GetPin_NotFound_ReturnsNull()
        {
            // Arrange
            var pinMapSheet = new PinMapSheet("PinMap");

            // Act
            Pin retrievedPin = pinMapSheet.GetPin("NonExistentPin");

            // Assert
            Assert.IsNull(retrievedPin);
        }

        [TestMethod]
        public void PinMapSheet_GetPinGroup_ByName()
        {
            // Arrange
            var pinMapSheet = new PinMapSheet("PinMap");
            var group = new PinGroup("Group1");
            var pin = new Pin { PinName = "Pin1" };
            group.AddPin(pin);
            pinMapSheet.AddGroup(group);

            // Act
            PinGroup retrievedGroup = pinMapSheet.GetGroup("Group1");

            // Assert
            Assert.IsNotNull(retrievedGroup);
            Assert.AreEqual("Group1", retrievedGroup.PinName);
        }

        [TestMethod]
        public void PinMapSheet_RemovePin()
        {
            // Arrange
            var pinMapSheet = new PinMapSheet("PinMap");
            var pin = new Pin { PinName = "Pin1" };
            pinMapSheet.AddPin(pin);

            // Act
            pinMapSheet.RemovePinAt(0);

            // Assert
            Assert.AreEqual(0, pinMapSheet.PinList.Count);
        }

        [TestMethod]
        public void PinMapSheet_SheetType_IsCorrect()
        {
            // Arrange & Act
            var pinMapSheet = new PinMapSheet("PinMap");

            // Assert
            Assert.AreEqual("DTPinMap", pinMapSheet.SheetType);
        }

        [TestMethod]
        public void PinMapSheet_Name_CanBeSet()
        {
            // Arrange
            var pinMapSheet = new PinMapSheet("PinMap")
            {
                // Act
                Name = "NewPinMapName"
            };

            // Assert
            Assert.AreEqual("NewPinMapName", pinMapSheet.Name);
        }

        [TestMethod]
        public void PinMapSheet_Write_WithEmptyPins_DoesNotThrow()
        {
            // Arrange
            var pinMapSheet = new PinMapSheet("PinMap");
            string tempFileName = Path.Combine(Path.GetTempPath(), $"PinMap_{Guid.NewGuid()}.txt");

            try
            {
                // Act
                pinMapSheet.Write(tempFileName);

                // Assert (no exception thrown is success)
                Assert.IsFalse(File.Exists(tempFileName));
            }
            finally
            {
                if (File.Exists(tempFileName))
                {
                    File.Delete(tempFileName);
                }
            }
        }

        [TestMethod]
        public void PinMapSheet_Write_WithVersion()
        {
            // Arrange
            var pinMapSheet = new PinMapSheet("PinMap");
            string tempFileName = Path.Combine(Path.GetTempPath(), $"PinMap_{Guid.NewGuid()}.txt");

            try
            {
                // Act
                pinMapSheet.Write(tempFileName, "2.0");

                // Assert
                Assert.IsFalse(File.Exists(tempFileName));
            }
            finally
            {
                if (File.Exists(tempFileName))
                {
                    File.Delete(tempFileName);
                }
            }
        }

        [TestMethod]
        public void PinMapSheet_AddMultiplePins_InSequence()
        {
            // Arrange
            var pinMapSheet = new PinMapSheet("PinMap");

            // Act
            for (int i = 0; i < 10; i++)
            {
                pinMapSheet.AddPin(new Pin { PinName = $"Pin{i}" });
            }

            // Assert
            Assert.AreEqual(10, pinMapSheet.PinList.Count);
            Assert.AreEqual("Pin0", pinMapSheet.PinList[0].PinName);
            Assert.AreEqual("Pin9", pinMapSheet.PinList[9].PinName);
        }

        [TestMethod]
        public void PinMapSheet_AddMultipleGroups_InSequence()
        {
            // Arrange
            var pinMapSheet = new PinMapSheet("PinMap");

            // Act
            for (int i = 0; i < 3; i++)
            {
                var group = new PinGroup($"Group{i}");
                var pin = new Pin { PinName = $"Pin{i}" };
                group.AddPin(pin);
                pinMapSheet.AddGroup(group);
            }

            // Assert
            Assert.AreEqual(3, pinMapSheet.GroupList.Count);
        }

        [TestMethod]
        public void PinMapSheet_GetPin_WithMultiplePins()
        {
            // Arrange
            var pinMapSheet = new PinMapSheet("PinMap");
            for (int i = 0; i < 5; i++)
            {
                pinMapSheet.AddPin(new Pin { PinName = $"Pin{i}" });
            }

            // Act
            Pin retrievedPin = pinMapSheet.GetPin("Pin3");

            // Assert
            Assert.IsNotNull(retrievedPin);
            Assert.AreEqual("Pin3", retrievedPin.PinName);
        }

        [TestMethod]
        public void PinMap_GetPinsFromGroup_ReturnsEmpty_WhenGroupDoesNotExist()
        {
            // Arrange
            var sheet = new PinMapSheet("");
            // Ensure internal dictionary (_groupDic) doesn't contain "UNKNOWN_GROUP"

            // Act
            List<Pin> result = sheet.GetPinsFromGroup("UNKNOWN_GROUP");

            // Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(0, result.Count);
        }

        [TestMethod]
        public void PinMap_GetPinsFromGroup_ResolvesNestedGroups_And_FallbackSinglePins()
        {
            // Arrange
            var sheet = new PinMapSheet("");

            // Setup target group structure:
            // "POWER_GRP" contains -> [Pin: "POWER_GRP" (Self reference loop breaker), Pin: "SUB_GRP", Pin: "ALONE_PIN", Pin: "MISSING_PIN"]
            var powerGroup = new PinGroup("POWER_GRP", PinMapConst.TypePower);
            powerGroup.AddPin(new Pin("POWER_GRP", "", ""));
            powerGroup.AddPin(new Pin("SUB_GRP", "", ""));
            powerGroup.AddPin(new Pin("ALONE_PIN", "", ""));
            powerGroup.AddPin(new Pin("MISSING_PIN", "", ""));

            // "SUB_GRP" contains -> [Pin: "VCC_CORE"]
            var subGroup = new PinGroup("SUB_GRP", PinMapConst.TypePower);
            subGroup.AddPin(new Pin("VCC_CORE", "", ""));

            sheet.AddGroup(powerGroup);
            sheet.AddGroup(subGroup);

            // Act
            List<Pin> result = sheet.GetPinsFromGroup("POWER_GRP");

            // Assert
            Assert.AreEqual(4, result.Count);
            Assert.AreEqual("POWER_GRP", result[0].PinName);
            Assert.AreEqual("VCC_CORE", result[1].PinName);
            Assert.AreEqual("ALONE_PIN", result[2].PinName);
            Assert.AreEqual("Does not exist in Pin map sheet", result[3].Comment);
        }

        [TestMethod]
        public void PinMap_DecompGroups_FlattensNestedGroupsIntoRawStrings()
        {
            // Arrange
            var sheet = new PinMapSheet("");

            // Mock TryGetGroup chain
            var mainGroup = new PinGroup("MAIN_GRP");
            mainGroup.AddPin(new Pin("MAIN_GRP", "I/O"));
            mainGroup.AddPin(new Pin("NESTED_GRP", "I/O"));

            var nestedGroup = new PinGroup("NESTED_GRP");
            nestedGroup.AddPin(new Pin("PIN_A", "I/O"));
            nestedGroup.AddPin(new Pin("PIN_B", "I/O"));

            sheet.AddGroup(mainGroup);
            sheet.AddGroup(nestedGroup);

            // Act
            List<string> result = sheet.DecompGroups("MAIN_GRP");

            // Assert
            Assert.AreEqual(3, result.Count);
            Assert.AreEqual("MAIN_GRP", result[0]);
            Assert.AreEqual("PIN_A", result[1]);
            Assert.AreEqual("PIN_B", result[2]);
        }

        [TestMethod]
        public void PinMap_DecompGroups_ReturnsInputString_WhenNoMatchingGroupFound()
        {
            // Arrange
            var sheet = new PinMapSheet("");

            // Act
            List<string> result = sheet.DecompGroups("ISOLATED_PIN");

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("ISOLATED_PIN", result[0]);
        }

        [TestMethod]
        public void PinMap_GenDcviGroup_FiltersAndSavesValidPowerAndAnalogGroups()
        {
            // Arrange
            var sheet = new PinMapSheet("");
            sheet.AddPin(new Pin { PinName = "VDD_1", ChannelType = "DCVI_01", PinType = PinMapConst.TypePower });
            sheet.AddPin(new Pin { PinName = "VREF_A", ChannelType = "DCVI_02", PinType = "Analog" });
            sheet.AddPin(new Pin { PinName = "SIG_OUT", ChannelType = "DIGITAL", PinType = "I/O" });

            // Act
            List<PinGroup> resultGroups = sheet.GenDcviGroup();

            // Assert
            Assert.AreEqual(2, resultGroups.Count);

            Assert.AreEqual("All_DCVI", resultGroups[0].PinName);
            Assert.AreEqual("VDD_1", resultGroups[0].PinList[0].PinName);

            Assert.AreEqual("All_DCVI_" + PinMapConst.TypeAnalog, resultGroups[1].PinName);
            Assert.AreEqual("VREF_A", resultGroups[1].PinList[0].PinName);
        }

        [TestMethod]
        public void PinMap_GenDcvsGroup_SetsDefaultColumnPropertiesAndSavesGroup()
        {
            // Arrange
            var sheet = new PinMapSheet("");
            sheet.AddPin(new Pin { PinName = "VDD_VS", ChannelType = "DCVS_A", PinType = PinMapConst.TypePower });

            // Act
            sheet.GenDcvsGroup();

            // Assert
            // Verify group registration side-effect occurred
            Assert.IsTrue(sheet.IsGroupExist("All_DCVS"));

            // Verify specific mutation payload
            bool found = sheet.TryGetGroup("All_DCVS", out PinGroup autogenGroup);
            Assert.IsTrue(found);
            Assert.AreEqual("Autogen Default", autogenGroup.PinList[0].ColumnA);
        }

        [TestMethod]
        public void PinMap_GetDiffGroupName_ReturnsGroupName_WhenExactPairMatches()
        {
            // Arrange
            var sheet = new PinMapSheet("");
            string[] pair = ["PIN_D_P", "PIN_D_N"];
            var matchingGroup = new PinGroup("USB_DIFF");
            matchingGroup.PinList.Add(new Pin("PIN_D_P", "I/O"));
            matchingGroup.PinList.Add(new Pin("PIN_D_N", "I/O"));
            sheet.AddGroup(matchingGroup);

            // Act
            string result = sheet.GetDiffGroupName(pair);

            // Assert
            Assert.AreEqual("USB_DIFF", result);
        }

        [TestMethod]
        public void PinMap_GetDiffGroupName_ReturnsEmpty_WhenPairLengthIsInvalid()
        {
            // Arrange
            var sheet = new PinMapSheet("");
            string[] invalidPair = ["PIN_A"];

            // Act
            string result = sheet.GetDiffGroupName(invalidPair);

            // Assert
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void PinMap_GetUartPinDic_PopulatesExistingUartGroups()
        {
            // Arrange
            var sheet = new PinMapSheet("");

            // Set up local state for simulated group registration checks
            sheet.AddGroup(new PinGroup("UART_TX") { PinList = { new Pin("PA_09", "I/O") } });
            sheet.AddGroup(new PinGroup("UART_RX") { PinList = { new Pin("PA_10", "I/O") } });
            sheet.AddPin(new Pin("PA_09", "I/O"));
            sheet.AddPin(new Pin("PA_10", "I/O"));

            // Act
            Dictionary<string, string> result = sheet.GetUartPinDic();

            // Assert
            Assert.AreEqual(2, result.Count);
            Assert.AreEqual("PA_09", result["UART_TX"]);
            Assert.AreEqual("PA_10", result["UART_RX"]);
        }

        [TestMethod]
        public void PinMap_GetPowerPins_FiltersOnlyPowerTypePins()
        {
            // Arrange
            var sheet = new PinMapSheet("");
            sheet.AddPin(new Pin { PinName = "VDD_CORE", PinType = PinMapConst.TypePower });
            sheet.AddPin(new Pin { PinName = "GPIO_01", PinType = PinMapConst.TypeIo });

            // Act
            List<Pin> result = sheet.GetPowerPins();

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("VDD_CORE", result[0].PinName);
        }

        [TestMethod]
        public void PinMap_GetIoPins_WithPatternPins_ExtractsMatchingIoPins()
        {
            // Arrange
            var sheet = new PinMapSheet("");
            sheet.AddPin(new Pin("PA_01", "I/O") { PinType = PinMapConst.TypeIo });
            sheet.AddPin(new Pin("VDD_A", "I/O") { PinType = PinMapConst.TypePower });
            var searchPatterns = new List<string> { "PA_01", "VDD_A", "UNKNOWN_PIN" };

            // Act
            List<string> result = sheet.GetIoPins(searchPatterns);

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("PA_01", result[0]);
        }

        [TestMethod]
        public void PinMap_GetIoPins_Overload_FiltersAllIoTypePins()
        {
            // Arrange
            var sheet = new PinMapSheet("");
            sheet.AddPin(new Pin { PinName = "PA_02", PinType = PinMapConst.TypeIo });
            sheet.AddPin(new Pin { PinName = "VDD_B", PinType = PinMapConst.TypePower });

            // Act
            List<Pin> result = sheet.GetIoPins();

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("PA_02", result[0].PinName);
        }

        [TestMethod]
        public void PinMap_GetIoContinuityPins_ExcludesSpecificNamingPatterns()
        {
            // Arrange
            var sheet = new PinMapSheet("");

            // ChannelType matches TypeIo to feed into the internal dependency method chain
            sheet.AddPin(new Pin { PinName = "DIG_IO_01", ChannelType = PinMapConst.TypeIo });
            sheet.AddPin(new Pin { PinName = "REFCLK_P", ChannelType = PinMapConst.TypeIo });
            sheet.AddPin(new Pin { PinName = "TEST_PA", ChannelType = PinMapConst.TypeIo });

            // Act
            List<Pin> result = sheet.GetIoContinuityPins();

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("DIG_IO_01", result[0].PinName);
        }

        [TestMethod]
        public void PinMap_GetAllDigitalDisconnectContinuityPins_ExcludesSenseAndMonitorPins()
        {
            // Arrange
            var sheet = new PinMapSheet("");
            sheet.AddPin(new Pin { PinName = "VDD_CORE", ChannelType = PinMapConst.TypeIo });
            sheet.AddPin(new Pin { PinName = "VDD_SENSE_01", ChannelType = PinMapConst.TypeIo });
            sheet.AddPin(new Pin { PinName = "VSS_SENSE_A", ChannelType = PinMapConst.TypeIo });
            sheet.AddPin(new Pin { PinName = "VDD_MONITOR", ChannelType = PinMapConst.TypeIo });

            // Act
            List<Pin> result = sheet.GetAllDigitalDisconnectContinuityPins();

            // Assert
            Assert.AreEqual(1, result.Count);
            Assert.AreEqual("VDD_CORE", result[0].PinName);
        }

        [TestMethod]
        public void PinMap_InsertPinAt_RegistersPinInListAndDictionary()
        {
            // Arrange
            var sheet = new PinMapSheet("");
            sheet.AddPin(new Pin("PA_00", "I/O"));
            sheet.AddPin(new Pin("PA_02", "I/O"));
            var targetPin = new Pin("PA_01", "I/O");

            // Act
            sheet.InsertPinAt(1, targetPin);

            // Assert
            Assert.AreEqual(3, sheet.PinList.Count);
            Assert.AreSame(targetPin, sheet.PinList[1]);
            Assert.AreSame(targetPin, sheet.GetPin("PA_01"));
        }

        [TestMethod]
        public void PinMap_InsertGroup_ThrowsException_WhenPinListIsEmpty()
        {
            // Arrange
            var sheet = new PinMapSheet("");
            var emptyGroup = new PinGroup("EMPTY_GRP");

            // Act
            Assert.ThrowsException<Exception>(() => sheet.InsertGroup(0, emptyGroup));
        }

        [TestMethod]
        public void PinMap_InsertGroup_AddsToCollectionAndDictionary_WhenGroupIsValid()
        {
            // Arrange
            var sheet = new PinMapSheet("");
            var validGroup = new PinGroup("VALID_GRP");
            validGroup.PinList.Add(new Pin("PA_01", "I/O"));

            // Act
            sheet.InsertGroup(0, validGroup);

            // Assert
            Assert.AreEqual(1, sheet.GroupList.Count);
            Assert.AreSame(validGroup, sheet.GroupList[0]);
            Assert.AreSame(validGroup, sheet.GetGroup("VALID_GRP"));
        }

        [TestMethod]
        public void PinMap_RemoveGroupAt_ThrowsException_WhenIndexIsOutOfBounds()
        {
            // Arrange
            var sheet = new PinMapSheet("");

            // Act
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => sheet.RemoveGroupAt(0));
        }

        [TestMethod]
        public void PinMap_RemoveGroupAt_RemovesFromCollectionAndDictionary_WhenIndexIsValid()
        {
            // Arrange
            var sheet = new PinMapSheet("");
            var targetGroup = new PinGroup("REMOVE_GRP");
            targetGroup.PinList.Add(new Pin("PA_01", "I/O"));

            sheet.AddGroup(targetGroup);

            // Act
            sheet.RemoveGroupAt(0);

            // Assert
            Assert.AreEqual(0, sheet.GroupList.Count);
            Assert.IsFalse(sheet.TryGetGroup("REMOVE_GRP", out _));
        }

        [TestMethod]
        public void PinMap_GetPinType_ReturnsType_WhenPinExists()
        {
            // Arrange
            var sheet = new PinMapSheet("");
            sheet.AddPin(new Pin("PA_02", "I/O") { PinType = PinMapConst.TypeIo });

            // Act
            string result = sheet.GetPinType("PA_02");

            // Assert
            Assert.AreEqual(PinMapConst.TypeIo, result);
        }

        [TestMethod]
        public void PinMap_GetPinType_ReturnsEmptyString_WhenPinDoesNotExist()
        {
            // Arrange
            var sheet = new PinMapSheet("");

            // Act
            string result = sheet.GetPinType("UNKNOWN_PIN");

            // Assert
            Assert.AreEqual("", result);
        }

        [TestMethod]
        public void PinMap_AddGroups_IteratesAndAddsAllProvidedGroups()
        {
            // Arrange
            var sheet = new PinMapSheet("");
            var group1 = new PinGroup("GRP_1") { PinList = { new Pin("PA_03", "I/O") } };
            var group2 = new PinGroup("GRP_2") { PinList = { new Pin("PA_04", "I/O") } };
            var groupsList = new List<PinGroup> { group1, group2 };

            // Act
            sheet.AddGroups(groupsList);

            // Assert
            // Verifies the downstream AddGroup mechanism was safely hit for each element
            Assert.IsTrue(sheet.IsGroupExist("GRP_1"));
            Assert.IsTrue(sheet.IsGroupExist("GRP_2"));
        }

        [TestMethod]
        public void PinMap_IsPinExist_ReturnsTrue_WhenSinglePinExists()
        {
            // Arrange
            var sheet = new PinMapSheet("");
            sheet.AddPin(new Pin("PA_01", "I/O"));

            // Act
            bool result = sheet.IsPinExist("PA_01");

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void PinMap_IsPinExist_ReturnsTrue_WhenGroupExists()
        {
            // Arrange
            var sheet = new PinMapSheet("");
            sheet.AddGroup(new PinGroup("GRP_POWER") { PinList = { new Pin("VDD", "I/O") } });

            // Act
            bool result = sheet.IsPinExist("GRP_POWER");

            // Assert
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void PinMap_IsPinExist_ReturnsFalse_WhenNeitherPinNorGroupExists()
        {
            // Arrange
            var sheet = new PinMapSheet("");

            // Act
            bool result = sheet.IsPinExist("UNKNOWN_IDENTIFIER");

            // Assert
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void PinMap_TryGetPin_ReturnsTrueAndOutputsPin_WhenPinExists()
        {
            // Arrange
            var sheet = new PinMapSheet("");
            var targetPin = new Pin("PA_02", "I/O") { PinType = PinMapConst.TypeIo };
            sheet.AddPin(targetPin);

            // Act
            bool found = sheet.TryGetPin("PA_02", out Pin resultPin);

            // Assert
            Assert.IsTrue(found);
            Assert.IsNotNull(resultPin);
            Assert.AreSame(targetPin, resultPin);
        }

        [TestMethod]
        public void PinMap_TryGetPin_ReturnsFalseAndOutputsNull_WhenPinDoesNotExist()
        {
            // Arrange
            var sheet = new PinMapSheet("");

            // Act
            bool found = sheet.TryGetPin("MISSING_PIN", out Pin resultPin);

            // Assert
            Assert.IsFalse(found);
            Assert.IsNull(resultPin);
        }
    }
}
