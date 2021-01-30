using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Drawing;
using System.IO;

namespace SW5E_Converter
{
    class Program
    {
        public static string JSON_FILE_NAME = "gil.json";
        public static string EXCEL_FILE_NAME = "test.xlsx";
        static void Main(string[] args)
        {
            using StreamReader r = new StreamReader(JSON_FILE_NAME);
            string json = r.ReadToEnd();
            Rootobject jsonObject = JsonConvert.DeserializeObject<Rootobject>(json);

            var fileInfo = new FileInfo(EXCEL_FILE_NAME);
            if (File.Exists(EXCEL_FILE_NAME))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var excelPackage = new OfficeOpenXml.ExcelPackage(fileInfo))
                {
                    ExcelWorksheet combatSheet = excelPackage.Workbook.Worksheets[1];
                    ExcelWorksheet characterSheet = excelPackage.Workbook.Worksheets[2];
                    ExcelWorksheet inventorySheet = excelPackage.Workbook.Worksheets[3];
                    ExcelWorksheet powersSheet = excelPackage.Workbook.Worksheets[4];

                    writeToCombatSheet(combatSheet, jsonObject);
                    //writeToCharacterSheet(characterSheet, jsonObject);
                    //writeToInventorySheet(inventorySheet, jsonObject);
                    //writeToPowersSheet(powersSheet, jsonObject);
                    excelPackage.Save();
                }
                Console.WriteLine(jsonObject.classes[0].name);
            }
        }

        private static void writeToPowersSheet(ExcelWorksheet powersSheet, Rootobject rootobject)
        {
            throw new NotImplementedException();
        }

        private static void writeToInventorySheet(ExcelWorksheet inventorySheet, Rootobject rootobject)
        {
            throw new NotImplementedException();
        }

        private static void writeToCharacterSheet(ExcelWorksheet characterSheet, Rootobject rootobject)
        {
            
        }

        private static void writeToCombatSheet(ExcelWorksheet combatSheet, Rootobject rootobject)
        {
            var cells = combatSheet.Cells;
            cells[5, 2].Value = rootobject.name;
            for (int i = 0; i < rootobject.classes.Length && i <= 4; i++)
            {
                writeString(cells[3 + i, 19], rootobject.classes[i].name);
                cells[3 + i, 26].Value = rootobject.classes[i].levels;
                cells[3 + i, 28].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                cells[3 + i, 32].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
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
