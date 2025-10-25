using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using Microsoft.VisualStudio.Shell;

namespace VSESwitchEnv
{
    internal static class Config
    {
        private static string _stateFile = null;

        public static bool Load()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string[] candidates = { ".vseswitchenv", ".editorconfig" };
            var configFile = candidates.Select(name => new FileInfo(Path.Combine(EnvSolution.SolDir, name))).FirstOrDefault(f => f.Exists);
            if (configFile == null)
                return false;

            string envName = null;
            string[] projNames = [];
            foreach (var str in File.ReadAllLines(configFile.FullName))
            {
                var line = str.Trim();
                if (line == string.Empty || line[0] == '#')
                    continue;

                if (line.Trim().StartsWith("["))
                {
                    if (line.StartsWith("[env:") && line.EndsWith("]"))
                    {
                        var parts = line.Replace("[env:", "").Replace("]", "").Split(new char[] { '|' }, 2);
                        envName = parts[0].Trim();
                        projNames = parts.Length == 2 ? parts[1].Split(',') : [Consts.SharedName];
                        continue;
                    }
                    envName = null;
                }
                else if (envName != null)
                {
                    var parts = line.Split(new char[] { '=' }, 2);
                    if (parts.Length == 2)
                    {
                        var envVar = new EnvVar(parts[0], parts[1]);
                        if (!envVar)
                        {
                            OutputPane.Write("Invalid format at: " + line);
                            continue;
                        }
                        foreach (var name in projNames)
                        {
                            var pName = name.Trim();
                            if (!EnvSolution.Projects.TryGetValue(pName, out var project))
                                EnvSolution.Projects[pName] = project = new(pName);
                            project.PushData(envName, envVar);
                        }
                    }
                }
            }
            return true;
        }

        public static void LoadStates()
        {
            _stateFile = Path.Combine(EnvSolution.ExtDir, "state.json");
            if (!File.Exists(_stateFile))
                return;

            var data = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(_stateFile));
            foreach (var v in data!)
            {
                if (EnvSolution.Projects.TryGetValue(v.Key, out var project))
                    EnvSolution.UpdateProps(project, v.Value);
            }
        }

        public static void SaveStates()
        {
            var states = new Dictionary<string, string>();
            foreach (var p in EnvSolution.Projects)
                states.Add(p.Key, p.Value.Selected);

            var options = new JsonSerializerOptions { WriteIndented = true };
            try
            {
                File.WriteAllText(_stateFile, JsonSerializer.Serialize(states, options));
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception during write state.json: {ex}");
            }
        }
    }
}