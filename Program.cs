using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace SW5E_Converter
{
    class Program
    {
        public static string JSON_FILE_NAME = "sw5e.json";
        public static string EXCEL_FILE_NAME = "OMNFormattedCharacterSheet.xlsx";
        static void Main(string[] args)
        {
            string jsonFilePath;
            string excelFilePath;
            if (!File.Exists(JSON_FILE_NAME) || !File.Exists(EXCEL_FILE_NAME))
            {
                Console.WriteLine("Enter the path to the SW5E .json file including filename");
                jsonFilePath = Console.ReadLine().Trim();
                while (!File.Exists(jsonFilePath))
                {
                    Console.WriteLine("Enter a valid path to the SW5E .json file including filename");
                    jsonFilePath = Console.ReadLine().Trim();
                }
                Console.WriteLine("Enter the path to the Ord Mantell Nights .xlsx file including filename");
                excelFilePath = Console.ReadLine().Trim();
                while (!File.Exists(jsonFilePath))
                {
                    Console.WriteLine("Enter a valid path to the Ord Mantell Nights .xlsx file including filename");
                    excelFilePath = Console.ReadLine().Trim();
                }

            } 
            else
            {
                jsonFilePath = JSON_FILE_NAME;
                excelFilePath = EXCEL_FILE_NAME;
            }
            using StreamReader r = new StreamReader(jsonFilePath);
            string json = r.ReadToEnd();
            Rootobject jsonObject = JsonConvert.DeserializeObject<Rootobject>(json);

            var fileInfo = new FileInfo(excelFilePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var excelPackage = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet combatSheet = excelPackage.Workbook.Worksheets[1];
                ExcelWorksheet characterSheet = excelPackage.Workbook.Worksheets[2];
                ExcelWorksheet inventorySheet = excelPackage.Workbook.Worksheets[3];
                ExcelWorksheet powersSheet = excelPackage.Workbook.Worksheets[4];
                ExcelWorksheet powerListSheet = excelPackage.Workbook.Worksheets[5];

                writeToCombatSheet(combatSheet.Cells, jsonObject);
                writeToCharacterSheet(characterSheet.Cells, jsonObject);
                writeToInventorySheet(inventorySheet.Cells, jsonObject);
                IList<Power> powerList = getPowerList(powerListSheet.Cells);
                writeToPowersSheet(powersSheet, jsonObject, powerList);
                excelPackage.Save();
            }
        }

        private static List<Power> getPowerList(ExcelRange cells)
        {
            var powerList = new List<Power>();
            for (int i = 4; i < 325; i++)
            {
                string alignment = cells[i, 4].Value.ToString();
                string name = cells[i, 3].Value.ToString();
                string castingPeriod = cells[i, 5].Value.ToString();
                string range = cells[i, 6].Value.ToString();
                string description = cells[i, 7].Value.ToString();
                string concentration = cells[i, 4].Value.ToString();
                string level = cells[i, 2].Value.ToString();

                powerList.Add(new Power(alignment, name, castingPeriod, range, description, concentration, level));
            }
            return powerList;
        }

        private static void writeToPowersSheet(ExcelWorksheet workSheet, Rootobject rootobject, IList<Power> powerList)
        {
            ExcelRange cells = workSheet.Cells;
            var knownForcePowerList = new List<string>();
            if (rootobject.classes[0].forcePowers != null)
            {
                knownForcePowerList.AddRange(rootobject.classes[0].forcePowers);
            }
            if (rootobject.customForcePowers != null)
            {
                knownForcePowerList.AddRange(rootobject.customForcePowers);
            }

            var knownTechPowerList = new List<string>();
            if (rootobject.classes[0].techPowers != null)
            {
                knownTechPowerList.AddRange(rootobject.classes[0].techPowers);
            }
            if (rootobject.customTechPowers != null)
            {
                knownTechPowerList.AddRange(rootobject.customTechPowers);
            }

            int level0Count = 0;
            int level1Count = 0;
            int level2Count = 0;

            foreach (string forcePower in knownForcePowerList)
            {
                Power power = powerList.FirstOrDefault(x => x.name.Equals(forcePower));
                if (power != null)
                {
                    if (power.level == "0")
                    {
                        if (level0Count >= 10)
                        {
                            workSheet.InsertRow(23 + level0Count, 1);
                        }
                        cells[23 + level0Count, 2].Value = power.alignment;
                        cells[23 + level0Count, 7].Value = power.name;
                        cells[23 + level0Count, 13].Value = power.castingPeriod;
                        cells[23 + level0Count, 19].Value = power.range;
                        cells[23 + level0Count, 23].Value = power.description;
                        cells[23 + level0Count, 53].Value = power.duration;
                        cells[23 + level0Count, 59].Value = power.concentration;
                        level0Count++;
                    }
                    else if (power.level == "1")
                    {
                        cells[35 + level0Count, 2].Value = power.alignment;
                        cells[35 + level0Count, 7].Value = power.name;
                        cells[35 + level0Count, 13].Value = power.castingPeriod;
                        cells[35 + level0Count, 19].Value = power.range;
                        cells[35 + level0Count, 23].Value = power.description;
                        cells[35 + level0Count, 53].Value = power.duration;
                        cells[35 + level0Count, 59].Value = power.concentration;
                        level1Count++;
                    }
                    else if (power.level == "2")
                    {
                        cells[45 + level0Count, 2].Value = power.alignment;
                        cells[45 + level0Count, 7].Value = power.name;
                        cells[45 + level0Count, 13].Value = power.castingPeriod;
                        cells[45 + level0Count, 19].Value = power.range;
                        cells[45 + level0Count, 23].Value = power.description;
                        cells[45 + level0Count, 53].Value = power.duration;
                        cells[45 + level0Count, 59].Value = power.concentration;
                        level2Count++;
                    }
                }
            }
            foreach (string techPower in knownTechPowerList)
            {
                Power power = powerList.FirstOrDefault(x => x.name == techPower);
                if (power != null)
                {
                    if (power.level == "0")
                    {
                        if (level0Count >= 10)
                        {
                            workSheet.InsertRow(23 + level0Count, 1);
                        }
                        cells[23 + level0Count, 2].Value = power.alignment;
                        cells[23 + level0Count, 7].Value = power.name;
                        cells[23 + level0Count, 13].Value = power.castingPeriod;
                        cells[23 + level0Count, 19].Value = power.range;
                        cells[23 + level0Count, 23].Value = power.description;
                        cells[23 + level0Count, 53].Value = power.duration;
                        cells[23 + level0Count, 59].Value = power.concentration;
                        level0Count++;
                    }
                    else if (power.level == "1")
                    {
                        cells[35 + level0Count, 2].Value = power.alignment;
                        cells[35 + level0Count, 7].Value = power.name;
                        cells[35 + level0Count, 13].Value = power.castingPeriod;
                        cells[35 + level0Count, 19].Value = power.range;
                        cells[35 + level0Count, 23].Value = power.description;
                        cells[35 + level0Count, 53].Value = power.duration;
                        cells[35 + level0Count, 59].Value = power.concentration;
                        level1Count++;
                    }
                    else if (power.level == "2")
                    {
                        cells[45 + level0Count, 2].Value = power.alignment;
                        cells[45 + level0Count, 7].Value = power.name;
                        cells[45 + level0Count, 13].Value = power.castingPeriod;
                        cells[45 + level0Count, 19].Value = power.range;
                        cells[45 + level0Count, 23].Value = power.description;
                        cells[45 + level0Count, 53].Value = power.duration;
                        cells[45 + level0Count, 59].Value = power.concentration;
                        level2Count++;
                    }
                }
            }
        }

        private static void writeToInventorySheet(ExcelRange cells, Rootobject rootobject)
        {
            Equipment[] equipmentArray = rootobject.equipment;
            Customequipment[] customEquipmentArray = rootobject.customEquipment;
            int count = 0;
            for (int i = 0; i < equipmentArray.Length && i < 50; i++)
            {
                Equipment equipment = equipmentArray[i];
                cells[3 + i, 2].Value = equipment.name + (" (insert cost)");
                cells[3 + i, 2].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                cells[3 + i, 19].Value = "";
                cells[3 + i, 19].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                cells[3 + i, 22].Value = equipment.quantity;

                cells[3 + i, 29].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                cells[3 + i, 30].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                cells[3 + i, 31].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                count++;
            }
            for (int i = 0; i < customEquipmentArray.Length && i < 50-count; i++)
            {
                Customequipment equipment = customEquipmentArray[i];
                if (equipment.quantity != 1)
                {
                    cells[3 + i + count, 2].Value = equipment.name + " (" + equipment.cost + " x " + equipment.quantity + ")";
                }
                else
                {
                    cells[3 + i + count, 2].Value = equipment.name + " (" + equipment.cost + ")";
                }
                cells[3 + i + count, 19].Value = equipment.weight;
                cells[3 + i + count, 22].Value = equipment.quantity;

                cells[3 + i + count, 29].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                cells[3 + i + count, 30].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                cells[3 + i + count, 31].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
            }

            cells[9, 35].Value = "Current Credits: " + rootobject.credits;
            cells[31, 35].Value = rootobject.notes;
        }

        private static void writeToCharacterSheet(ExcelRange cells, Rootobject rootobject)
        {
            Characteristics characteristics = rootobject.characteristics;
            cells[3, 26].Value = characteristics.PlaceofBirth;
            cells[6, 23].Value = characteristics.Age;
            cells[7, 23].Value = characteristics.Height;
            //size?     cells[8, 23].Value = characteristics.;
            cells[9, 23].Value = characteristics.Eyes;
            cells[6, 47].Value = characteristics.Gender;
            cells[7, 47].Value = characteristics.Weight;
            cells[8, 47].Value = characteristics.Hair;
            cells[9, 47].Value = characteristics.Skin;
            cells[10, 25].Value = characteristics.Appearance;
            cells[13, 27].Value = characteristics.PersonalityTraits;
            cells[17, 22].Value = characteristics.Ideal;
            cells[20, 22].Value = characteristics.Bond;
            cells[23, 22].Value = characteristics.Flaw;
            cells[42, 19].Value = characteristics.Backstory;

            cells[14, 2].Value = "Galactic Basic";
            string[] languages = rootobject.customLanguages;
            for (int i = 0; i < languages.Length; i++)
            {
                cells[15 + i, 2].Value = languages[i];
            }
        }

        private static void writeToCombatSheet(ExcelRange cells, Rootobject rootobject)
        {
            cells[5, 2].Value = rootobject.name;
            for (int i = 0; i < rootobject.classes.Length && i <= 4; i++)
            {
                writeString(cells[3 + i, 19], rootobject.classes[i].name);
                cells[3 + i, 26].Value = rootobject.classes[i].levels;
                cells[3 + i, 28].Value = rootobject.classes[i].getHitDice();
                cells[3 + i, 32].Value = rootobject.classes[i].getHP();
            }

            cells[1, 37].Value = rootobject.characteristics.alignment;
            cells[1, 43].Value = rootobject.species.name;
            cells[1, 56].Value = rootobject.background.name;
            cells[3, 37].Value = rootobject.experiencePoints;
            cells[3, 47].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
            if (string.IsNullOrWhiteSpace(rootobject.user))
            { 
                cells[3, 58].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
            }
            else 
            { 
                cells[3, 58].Value = rootobject.user; 
            }

            cells[10, 2].Value = rootobject.baseAbilityScores.Strength;
            cells[14, 2].Value = rootobject.baseAbilityScores.Dexterity;
            cells[20, 2].Value = rootobject.baseAbilityScores.Constitution;
            cells[24, 2].Value = rootobject.baseAbilityScores.Intelligence;
            cells[32, 2].Value = rootobject.baseAbilityScores.Wisdom;
            cells[40, 2].Value = rootobject.baseAbilityScores.Charisma;

            writeAbilityScores(cells, rootobject);
        }

        private static void writeString(ExcelRange cell, string writeObject)
        {
            if (string.IsNullOrWhiteSpace(writeObject))
            {
                cell.Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
            }
            else
            {
                cell.Value = writeObject;
            }
        }

        private static void writeAbilityScores(ExcelRange cells, Rootobject rootobject)
        {
            var abilityScores = rootobject.tweaks.abilityScores;
            cells[10, 2].Value = rootobject.baseAbilityScores.Strength + rootobject.species.abilityScoreImprovement.Strength;
            cells[14, 2].Value = rootobject.baseAbilityScores.Dexterity + rootobject.species.abilityScoreImprovement.Dexterity;
            cells[20, 2].Value = rootobject.baseAbilityScores.Constitution + rootobject.species.abilityScoreImprovement.Constitution;
            cells[24, 2].Value = rootobject.baseAbilityScores.Intelligence + +rootobject.species.abilityScoreImprovement.Intelligence;
            cells[32, 2].Value = rootobject.baseAbilityScores.Wisdom + rootobject.species.abilityScoreImprovement.Wisdom;
            cells[40, 2].Value = rootobject.baseAbilityScores.Charisma + rootobject.species.abilityScoreImprovement.Charisma;
            if (abilityScores != null)
            {
                Ability strength = abilityScores.Strength;
                if (strength != null)
                {
                    writeSavingThrow(cells, strength.savingThrowModifier, 9);
                    Skills skills = strength.skills;
                    writeSkill(cells, skills.Athletics, 10);
                }
                Ability dexterity = abilityScores.Dexterity;
                if (dexterity != null)
                {
                    writeSavingThrow(cells, dexterity.savingThrowModifier, 13);
                    Skills skills = dexterity.skills;
                    writeSkill(cells, skills.Acrobatics, 14);
                    writeSkill(cells, skills.SleightofHand, 15);
                    writeSkill(cells, skills.Stealth, 16);
                }
                Ability constitution = abilityScores.Constitution;
                if (constitution != null)
                {
                    writeSavingThrow(cells, constitution.savingThrowModifier, 19);
                }
                Ability intelligence = abilityScores.Intelligence;
                if (intelligence != null)
                {
                    writeSavingThrow(cells, intelligence.savingThrowModifier, 23);
                    Skills skills = intelligence.skills;
                    writeSkill(cells, skills.Investigation, 24);
                    writeSkill(cells, skills.Lore, 25);
                    writeSkill(cells, skills.Nature, 26);
                    writeSkill(cells, skills.Piloting, 27);
                    writeSkill(cells, skills.Technology, 28);
                }
                Ability wisdom = abilityScores.Wisdom;
                if (wisdom != null)
                {
                    writeSavingThrow(cells, wisdom.savingThrowModifier, 31);
                    Skills skills = wisdom.skills;
                    writeSkill(cells, skills.AnimalHandling, 32);
                    writeSkill(cells, skills.Insight, 33);
                    writeSkill(cells, skills.Medicine, 34);
                    writeSkill(cells, skills.Perception, 35);
                    writeSkill(cells, skills.Survival, 36);
                }
                Ability charisma = abilityScores.Charisma;
                if (charisma != null)
                {
                    writeSavingThrow(cells, charisma.savingThrowModifier, 39);
                    Skills skills = charisma.skills;
                    writeSkill(cells, skills.Deception, 40);
                    writeSkill(cells, skills.Intimidation, 41);
                    writeSkill(cells, skills.Performance, 42);
                    writeSkill(cells, skills.Persuasion, 43);
                }
            }
        }

        static void writeSavingThrow(ExcelRange cells, Savingthrowmodifier savingthrowmodifier, int cellStart)
        {
            if (savingthrowmodifier != null)
            {
                if (savingthrowmodifier.proficiency == "Proficient")
                {
                    cells[cellStart, 11].Value = "•";
                }
                if (savingthrowmodifier.bonus != null)
                {
                    cells[cellStart, 13].Value = savingthrowmodifier.bonus;
                }
            }
        }

        static void writeSkill(ExcelRange cells, Skill skill, int cellStart)
        {
            if (skill != null)
            {
                if (skill.proficiency == "Proficient")
                {
                    cells[cellStart, 11].Value = "●";
                }
                else if (skill.proficiency == "Expertise")
                {
                    cells[cellStart, 11].Value = "x2";
                }
                if (skill.bonus != null)
                {
                    cells[cellStart, 13].Value = skill.bonus;
                }
            }
        }
    }
}
