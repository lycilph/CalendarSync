using System;
using System.IO;
using Newtonsoft.Json;

namespace CalendarSync.Data
{
    public class Settings
    {
        public static readonly int StartupDelayInSeconds = 30;
        public static readonly int MinSyncIntervalInMinutes = 5;
        public static readonly string ApplicationName = "CalenderSync";
        public static readonly string DataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), ApplicationName);
        public static readonly string SecretsFile = Path.Combine(DataDir, "ClientSecrets.json");
        public static readonly string SettingsFile = Path.Combine(DataDir, "Settings.json");

        public bool IsSyncEnabled { get; set; } = false;
        public int SyncWindowInMonths { get; set; } = 1;
        public int SyncIntervalInMinutes { get; set; } = 10;
        public Calendar Calendar { get; set; }

        public static Settings Load()
        {
            if (!Directory.Exists(DataDir))
            {
                Directory.CreateDirectory(DataDir);
            }

            var empty = new Settings();
            if (!File.Exists(SettingsFile))
            {
                empty.Save();
                return empty;
            }

            var json = File.ReadAllText(SettingsFile);
            return JsonConvert.DeserializeObject<Settings>(json);
        }

        public void Save()
        {
            var json = JsonConvert.SerializeObject(this);
            File.WriteAllText(SettingsFile, json);
        }
    }
}