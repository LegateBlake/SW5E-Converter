using Newtonsoft.Json;

namespace SW5E_Converter
{

    public class Rootobject
    {
        public string name { get; set; }
        public string id { get; set; }
        public string userId { get; set; }
        public string builderVersion { get; set; }
        public string image { get; set; }
        public string user { get; set; }
        public int experiencePoints { get; set; }
        public Species species { get; set; }
        public Class1[] classes { get; set; }
        public Baseabilityscores baseAbilityScores { get; set; }
        public Background background { get; set; }
        public Characteristics characteristics { get; set; }
        public int credits { get; set; }
        public Equipment[] equipment { get; set; }
        public Currentstats currentStats { get; set; }
        public Tweaks tweaks { get; set; }
        public Customproficiency[] customProficiencies { get; set; }
        public string[] customLanguages { get; set; }
        public Customfeature[] customFeatures { get; set; }
        public string[] customFeats { get; set; }
        public string[] customTechPowers { get; set; }
        public object[] customForcePowers { get; set; }
        public Customequipment[] customEquipment { get; set; }
        public Settings settings { get; set; }
        public string notes { get; set; }
        public long createdAt { get; set; }
        public long changedAt { get; set; }
        public string localId { get; set; }
    }

    public class Species
    {
        public string name { get; set; }
        public int abilityScoreImprovementSelectedOption { get; set; }
        public Abilityscoreimprovement abilityScoreImprovement { get; set; }
    }

    public class Abilityscoreimprovement
    {
        public int Strength { get; set; } 
        public int Dexterity { get; set; } 
        public int Constitution { get; set; } 
        public int Intelligence { get; set; } 
        public int Wisdom { get; set; } 
        public int Charisma { get; set; } 
    }

    public class Baseabilityscores
    {
        public int Strength { get; set; }
        public int Dexterity { get; set; }
        public int Constitution { get; set; }
        public int Intelligence { get; set; }
        public int Wisdom { get; set; }
        public int Charisma { get; set; }
    }

    public class Background
    {
        public string name { get; set; }
        public Feat feat { get; set; }
        public string feature { get; set; }
    }

    public class Feat
    {
        public string name { get; set; }
        public string type { get; set; }
    }

    public class Characteristics
    {
        public string alignment { get; set; }
        public string PersonalityTraits { get; set; }
        public string Ideal { get; set; }
        public string Bond { get; set; }
        public string Flaw { get; set; }
        public string Gender { get; set; }
        public string PlaceofBirth { get; set; }
        public string Age { get; set; }
        public string Height { get; set; }
        public string Weight { get; set; }
        public string Hair { get; set; }
        public string Eyes { get; set; }
        public string Skin { get; set; }
        public string Appearance { get; set; }
        public string Backstory { get; set; }
    }

    public class Currentstats
    {
        public int hitPointsLost { get; set; }
        public int temporaryHitPoints { get; set; }
        public int techPointsUsed { get; set; }
        public int forcePointsUsed { get; set; }
        public int superiorityDiceUsed { get; set; }
        public Hitdiceused hitDiceUsed { get; set; }
        public Deathsaves deathSaves { get; set; }
        public bool hasInspiration { get; set; }
        public Featurestimesused featuresTimesUsed { get; set; }
        public object[] conditions { get; set; }
        public int exhaustion { get; set; }
        public Highlevelcasting highLevelCasting { get; set; }
    }

    public class Hitdiceused
    {
    }

    public class Deathsaves
    {
        public int successes { get; set; }
        public int failures { get; set; }
    }

    public class Featurestimesused
    {
    }

    public class Highlevelcasting
    {
        public bool level6 { get; set; }
        public bool level7 { get; set; }
        public bool level8 { get; set; }
        public bool level9 { get; set; }
    }

    public class Tweaks
    {
        public Abilityscores abilityScores { get; set; }
    }

    public class Abilityscores
    {
        public Ability Strength { get; set; }
        public Ability Dexterity { get; set; }
        public Ability Constitution { get; set; }
        public Ability Intelligence { get; set; }
        public Ability Wisdom { get; set; }
        public Ability Charisma { get; set; }
    }

    public class Ability
    {
        public Savingthrowmodifier savingThrowModifier { get; set; }
        public Skills skills { get; set; }
    }

    public class Savingthrowmodifier
    {
        public string proficiency { get; set; }
        public int _override { get; set; }
        public string bonus { get; set; }
    }

    public class Skills
    {
        public Skill Athletics { get; set; }
        public Skill Acrobatics { get; set; }
        [JsonProperty(PropertyName = "Sleight of Hand")]
        public Skill SleightofHand { get; set; }
        public Skill Stealth { get; set; }
        public Skill Investigation { get; set; }
        public Skill Lore { get; set; }
        public Skill Nature { get; set; }
        public Skill Piloting { get; set; }
        public Skill Technology { get; set; }
        [JsonProperty(PropertyName = "Animal Handling")]
        public Skill AnimalHandling { get; set; }
        public Skill Insight { get; set; }
        public Skill Medicine { get; set; }
        public Skill Perception { get; set; }
        public Skill Survival { get; set; }
        public Skill Deception { get; set; }
        public Skill Intimidation { get; set; }
        public Skill Performance { get; set; }
        public Skill Persuasion { get; set; }
    }

    public class Skill
    {
        public string bonus { get; set; }
        public int _override { get; set; }
        public string proficiency { get; set; }
    }

    public class Settings
    {
        public bool isEnforcingForcePrerequisites { get; set; }
        public bool isFixedHitPoints { get; set; }
        public string abilityScoreMethod { get; set; }
    }

    public class Class1
    {
        public string name { get; set; }
        public int levels { get; set; }
        public object[] hitPoints { get; set; }
        public object[] abilityScoreImprovements { get; set; }
        public string[] forcePowers { get; set; }
    }

    public class Equipment
    {
        public string name { get; set; }
        public int quantity { get; set; }
        public string category { get; set; }
    }

    public class Customproficiency
    {
        public string name { get; set; }
        public string type { get; set; }
        public string proficiencyLevel { get; set; }
    }

    public class Customfeature
    {
        public string name { get; set; }
        public string content { get; set; }
    }

    public class Customequipment
    {
        public string name { get; set; }
        public int quantity { get; set; }
        public string equipmentCategory { get; set; }
        public int cost { get; set; }
        public string description { get; set; }
        public int weight { get; set; }
    }
}