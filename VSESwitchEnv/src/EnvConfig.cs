using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace VSESwitchEnv
{
    public class EnvData()
    {
        public class Var(string name, string value)
        {
            public readonly string Name = name.Split(':')[0].Trim();
            public readonly string Value = value.Trim();
            public readonly bool Define = name.Contains(":");

            public static implicit operator bool(Var i) => !string.IsNullOrEmpty(i.Name);
        }

        private string _selected = null;
        private readonly Dictionary<string, List<Var>> _data = [];

        public bool IsEmpty() => _data.Count == 0;

        public bool Selected(string selected)
        {
            if (selected == null || !_data.ContainsKey(selected) || _selected == selected)
                return false;

            _selected = selected;
            return true;
        }

        public string Selected() { return _selected; }
        public Dictionary<string, List<Var>> List() { return _data; }
        public List<Var> Current() { return (_selected != null && _data.TryGetValue(_selected, out var data)) ? data : []; }

        public void Add(string key, Var var)
        {
            if (!_data.ContainsKey(key))
                _data.Add(key, []);
            _data[key].Add(var);
        }

        public void Clear()
        {
            _selected = null;
            _data.Clear();
        }
    }

    internal static class EnvConfig
    {
        private static string _stateFile = null;

        public static void Load()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string[] candidates = { ".vseswitchenv", ".editorconfig" };
            var configFile = candidates.Select(name => new FileInfo(Path.Combine(EnvSolution.SolDir, name))).FirstOrDefault(f => f.Exists);
            if (configFile == null)
                return;

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
                        var parts = line.Replace("[env:", "").Replace("]", "").Split(['|'], 2);
                        envName = parts[0].Trim();
                        projNames = parts.Length == 2 ? parts[1].Split(',') : [];
                        continue;
                    }
                    envName = null;
                }
                else if (envName != null)
                {
                    var parts = line.Split(['='], 2);
                    if (parts.Length == 2)
                    {
                        var envVar = new EnvData.Var(parts[0], parts[1]);
                        if (!envVar)
                        {
                            OutputPane.Write("Invalid format at: " + line);
                            continue;
                        }
                        if (projNames.Length > 0)
                        {
                            foreach (var name in projNames)
                            {
                                var pName = name.Trim();
                                if (!EnvSolution.Projects.TryGetValue(pName, out var project))
                                    EnvSolution.Projects[pName] = project = new(pName);
                                project.Add(envName, envVar);
                            }
                        }
                        else
                            EnvSolution.SharedData.Add(envName, envVar);
                    }
                }
            }

            _stateFile = Path.Combine(EnvSolution.ExtDir, "state.json");
            if (File.Exists(_stateFile))
            {
                var states = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(_stateFile));
                foreach (var state in states)
                {
                    if (state.Key == Consts.SharedName)
                        EnvSolution.SharedData.Selected(state.Value);
                    else if (EnvSolution.Projects.TryGetValue(state.Key, out var p))
                        p.Selected(state.Value);
                }
            }
        }

        public static void Save()
        {
            if (_stateFile == null)
                return;

            var states = new Dictionary<string, string>();
            foreach (var p in EnvSolution.Projects)
                states.Add(p.Key, p.Value.Selected());
            if (!EnvSolution.SharedData.IsEmpty())
                states.Add(Consts.SharedName, EnvSolution.SharedData.Selected());

            try
            {
                var options = new JsonSerializerOptions { WriteIndented = true };
                File.WriteAllText(_stateFile, JsonSerializer.Serialize(states, options));
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception during write state.json: {ex}");
            }
        }
    }
}
