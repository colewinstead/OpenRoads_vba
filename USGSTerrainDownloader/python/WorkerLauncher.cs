using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;

internal static class WorkerLauncher
{
    private static readonly Encoding Utf8 = new UTF8Encoding(false);

    private static string Text(Dictionary<string, object> values, string key)
    {
        object value;
        if (!values.TryGetValue(key, out value) || value == null) return "";
        if (!(value is string)) throw new ArgumentException(key + " must be a string.");
        return (string)value;
    }

    private static string QuotePath(string path)
    {
        if (path.IndexOfAny(new[] { '"', '\r', '\n' }) >= 0 || path.EndsWith("\\"))
            throw new ArgumentException("Invalid command path.");
        return "\"" + path + "\"";
    }

    private static int Main(string[] args)
    {
        string folder = null;
        try
        {
            if (args.Length != 1) throw new ArgumentException("Usage: launch_worker.exe <request.json>");
            string request = Path.GetFullPath(args[0]);
            folder = Path.GetDirectoryName(request);
            if (new FileInfo(request).Length > 4 * 1024 * 1024)
                throw new ArgumentException("Request exceeds 4 MiB.");
            var serializer = new JavaScriptSerializer { MaxJsonLength = 4 * 1024 * 1024 };
            var config = serializer.Deserialize<Dictionary<string, object>>(File.ReadAllText(request, Encoding.UTF8));
            if (config == null) throw new ArgumentException("Request must be a JSON object.");
            string python = Text(config, "python_executable");
            if (!Path.IsPathRooted(python) || !File.Exists(python))
                throw new FileNotFoundException("Python executable not found. Configure a full python.exe path.");
            if (!String.Equals(Path.GetExtension(python), ".exe", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Select python.exe, not a .bat/.cmd launcher.");
            string script = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "main.py");
            if (!File.Exists(script)) throw new FileNotFoundException("Keep launch_worker.exe beside main.py.");
            var start = new ProcessStartInfo(python, QuotePath(script) + " " + QuotePath(request)) {
                UseShellExecute = false, CreateNoWindow = true,
                RedirectStandardOutput = true, RedirectStandardError = true,
                WorkingDirectory = folder
            };
            string environment = Text(config, "environment_file");
            if (environment.Length > 0)
            {
                var values = serializer.Deserialize<Dictionary<string, object>>(File.ReadAllText(environment, Encoding.UTF8));
                if (values == null) throw new ArgumentException("Environment must be a JSON object.");
                foreach (string key in values.Keys)
                {
                    string value = Text(values, key);
                    switch (key)
                    {
                        case "PATH_prepend": start.EnvironmentVariables["PATH"] = value + ";" + start.EnvironmentVariables["PATH"]; break;
                        case "GDAL_DATA": case "PROJ_DATA": case "PYTHONHOME": case "PYTHONPATH":
                        case "QT_PLUGIN_PATH": case "SSL_CERT_FILE": start.EnvironmentVariables[key] = value; break;
                        default: throw new ArgumentException("Unsupported child environment key: " + key);
                    }
                }
            }
            start.EnvironmentVariables["PROJ_NETWORK"] = "OFF";
            start.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";
            start.StandardOutputEncoding = Encoding.UTF8;
            start.StandardErrorEncoding = Encoding.UTF8;
            using (var process = Process.Start(start))
            {
                var stdout = process.StandardOutput.ReadToEndAsync();
                var stderr = process.StandardError.ReadToEndAsync();
                process.WaitForExit();
                File.WriteAllText(Path.Combine(folder, "stdout.log"), stdout.Result, Utf8);
                File.WriteAllText(Path.Combine(folder, "stderr.log"), stderr.Result, Utf8);
                if (process.ExitCode != 0 && !File.Exists(Path.Combine(folder, "status.txt")))
                {
                    File.WriteAllText(Path.Combine(folder, "status.txt"), "failed\r\nPython did not initialize. Verify environment paths; see stderr.log.", Encoding.Unicode);
                    File.WriteAllText(Path.Combine(folder, "error.log"), stderr.Result, Utf8);
                }
                return process.ExitCode;
            }
        }
        catch (Exception error)
        {
            if (folder != null && Directory.Exists(folder))
            {
                try
                {
                    File.WriteAllText(Path.Combine(folder, "status.txt"), "failed\r\n" + error.Message, Encoding.Unicode);
                    File.WriteAllText(Path.Combine(folder, "error.log"), error.ToString(), Utf8);
                }
                catch (IOException) { } // A failed log write must not open a Windows crash dialog.
                catch (UnauthorizedAccessException) { }
            }
            return 2;
        }
    }
}
