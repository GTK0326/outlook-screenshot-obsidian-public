using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = BuildArguments(args),
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            using (var process = new Process { StartInfo = startInfo })
            {
                process.Start();
                var stdout = process.StandardOutput.ReadToEnd();
                var stderr = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (process.ExitCode != 0)
                {
                    var message = BuildErrorMessage(process.ExitCode, stdout, stderr);
                    MessageBox.Show(message, "outlook-screenshot", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                return process.ExitCode;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString(), "outlook-screenshot", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 1;
        }
    }

    private static string BuildArguments(IEnumerable<string> args)
    {
        var scriptPath = ResolveScriptPath();
        var tokens = new List<string>
        {
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            scriptPath
        };
        tokens.AddRange(args);
        return string.Join(" ", tokens.Select(QuoteArgument));
    }

    private static string ResolveScriptPath()
    {
        var configuredPath = Environment.GetEnvironmentVariable("OUTLOOK_SCREENSHOT_SCRIPT_PATH");
        if (!string.IsNullOrWhiteSpace(configuredPath))
        {
            return configuredPath.Trim();
        }

        return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Import-OutlookMeetingsFromScreenshot.ps1");
    }

    private static string BuildErrorMessage(int exitCode, string stdout, string stderr)
    {
        var builder = new StringBuilder();
        builder.AppendLine("outlook-screenshot failed.");
        builder.AppendLine("Exit code: " + exitCode);

        var details = string.IsNullOrWhiteSpace(stderr) ? stdout : stderr;
        if (!string.IsNullOrWhiteSpace(details))
        {
            builder.AppendLine();
            builder.AppendLine(TrimForDialog(details));
        }

        return builder.ToString().Trim();
    }

    private static string TrimForDialog(string text)
    {
        var normalized = text.Replace("\r\n", "\n").Trim();
        if (normalized.Length <= 1400)
        {
            return normalized;
        }

        return normalized.Substring(0, 1400) + "\n...";
    }

    private static string QuoteArgument(string arg)
    {
        if (string.IsNullOrEmpty(arg))
        {
            return "\"\"";
        }

        if (arg.IndexOfAny(new[] { ' ', '\t', '"' }) < 0)
        {
            return arg;
        }

        var builder = new StringBuilder();
        builder.Append('"');
        var backslashCount = 0;

        foreach (var ch in arg)
        {
            if (ch == '\\')
            {
                backslashCount++;
                continue;
            }

            if (ch == '"')
            {
                builder.Append('\\', backslashCount * 2 + 1);
                builder.Append('"');
                backslashCount = 0;
                continue;
            }

            if (backslashCount > 0)
            {
                builder.Append('\\', backslashCount);
                backslashCount = 0;
            }

            builder.Append(ch);
        }

        if (backslashCount > 0)
        {
            builder.Append('\\', backslashCount * 2);
        }

        builder.Append('"');
        return builder.ToString();
    }
}
