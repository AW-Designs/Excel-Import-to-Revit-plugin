using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelScheduleImporter
{
    /// <summary>
    /// Self-update against GitHub Releases.
    ///
    /// On Revit launch (background thread) this checks the repo's latest release.
    /// If the release version is newer than the running assembly, it downloads the
    /// zip asset that matches this Revit year, extracts it to a local staging
    /// folder, and arms a small detached process that waits for THIS Revit to
    /// close, then swaps the new files into the add-in folder. The next Revit
    /// launch runs the updated version.
    ///
    /// This design sidesteps the file lock: the running DLL is never overwritten
    /// while Revit holds it - the swap happens only once Revit has exited.
    /// </summary>
    public static class Updater
    {
        private const string Owner = "AW-Designs";
        private const string RepoName = "Excel-Import-to-Revit-plugin";
        private const string LatestApi =
            "https://api.github.com/repos/" + Owner + "/" + RepoName + "/releases/latest";

        // Only hit the GitHub API this often per machine (unauthenticated API is
        // 60 requests/hour per IP - an office behind one NAT must stay well under).
        private static readonly TimeSpan CheckInterval = TimeSpan.FromHours(4);

        /// <summary>Fire-and-forget update check. Never throws into Revit.</summary>
        public static void CheckInBackground(int revitYear, string liveDir)
        {
            if (revitYear == 0) return;
            Task.Run(() =>
            {
                try { Check(revitYear, liveDir); }
                catch (Exception ex) { App.LogCrash("Updater", ex); }
            });
        }

        private static string LocalRoot(int year)
        {
            string p = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ExcelScheduleImporter", year.ToString());
            Directory.CreateDirectory(p);
            return p;
        }

        private static void Check(int revitYear, string liveDir)
        {
            // Throttle: skip if we checked recently.
            string stamp = Path.Combine(LocalRoot(revitYear), "lastcheck.txt");
            if (File.Exists(stamp) &&
                DateTime.UtcNow - File.GetLastWriteTimeUtc(stamp) < CheckInterval)
                return;
            File.WriteAllText(stamp, DateTime.UtcNow.ToString("o"));

            // TLS 1.2 for older frameworks / GitHub's requirements.
            try { ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12; } catch { }

            string json = HttpGetString(LatestApi);
            if (string.IsNullOrEmpty(json)) return;

            Version latest = ParseVersion(Match(json, "\"tag_name\"\\s*:\\s*\"([^\"]+)\""));
            Version current = Assembly.GetExecutingAssembly().GetName().Version;
            if (latest == null || latest <= current) return;

            // Pick the asset whose name contains this Revit year, e.g. "...-R26.zip".
            string assetUrl = FindAssetUrl(json, "R" + revitYear);
            if (assetUrl == null) return;

            // Download + extract to a clean staging payload folder.
            string staging = Path.Combine(LocalRoot(revitYear), "staging");
            if (Directory.Exists(staging)) TryDeleteDir(staging);
            Directory.CreateDirectory(staging);

            string zipPath = Path.Combine(staging, "payload.zip");
            if (!DownloadFile(assetUrl, zipPath)) return;

            string payload = Path.Combine(staging, "payload");
            Directory.CreateDirectory(payload);
            ZipFile.ExtractToDirectory(zipPath, payload);
            try { File.Delete(zipPath); } catch { }

            // Arm the detached swapper: it waits for this Revit to exit, then
            // copies payload -> liveDir. Independent of OnShutdown, so it still
            // applies even if Revit is force-closed.
            ArmSwapper(payload, liveDir, revitYear, latest.ToString());
        }

        // ── HTTP ────────────────────────────────────────────────────────────

        private static string HttpGetString(string url)
        {
            using (var http = MakeClient())
                return http.GetStringAsync(url).GetAwaiter().GetResult();
        }

        private static bool DownloadFile(string url, string dest)
        {
            using (var http = MakeClient())
            using (var resp = http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead)
                                  .GetAwaiter().GetResult())
            {
                if (!resp.IsSuccessStatusCode) return false;
                using (var fs = new FileStream(dest, FileMode.Create, FileAccess.Write, FileShare.None))
                    resp.Content.CopyToAsync(fs).GetAwaiter().GetResult();
            }
            return true;
        }

        private static HttpClient MakeClient()
        {
            var http = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
            // GitHub API rejects requests without a User-Agent.
            http.DefaultRequestHeaders.UserAgent.ParseAdd("ExcelScheduleImporter-Updater");
            http.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
            return http;
        }

        // ── JSON (minimal, dependency-free) ─────────────────────────────────

        private static string Match(string s, string pattern)
        {
            var m = Regex.Match(s, pattern);
            return m.Success ? m.Groups[1].Value : null;
        }

        /// <summary>Find the browser_download_url of the .zip asset whose name contains <paramref name="marker"/>.</summary>
        private static string FindAssetUrl(string json, string marker)
        {
            foreach (Match m in Regex.Matches(json,
                "\"browser_download_url\"\\s*:\\s*\"([^\"]+\\.zip)\""))
            {
                string url = m.Groups[1].Value;
                if (url.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0)
                    return url;
            }
            return null;
        }

        private static Version ParseVersion(string tag)
        {
            if (string.IsNullOrWhiteSpace(tag)) return null;
            // strip a leading 'v' and anything non-numeric at the start
            var m = Regex.Match(tag, "(\\d+(?:\\.\\d+){0,3})");
            if (!m.Success) return null;
            var parts = m.Value.Split('.');
            int Get(int i) => i < parts.Length && int.TryParse(parts[i], out var v) ? v : 0;
            return new Version(Get(0), Get(1), Get(2), Get(3));
        }

        // ── Swapper ─────────────────────────────────────────────────────────

        private static void ArmSwapper(string payloadDir, string liveDir, int year, string version)
        {
            int pid = Process.GetCurrentProcess().Id;
            string script = Path.Combine(LocalRoot(year), "apply-update.ps1");

            // Waits for this Revit PID to exit, copies the payload into the live
            // add-in folder, records the applied version, then cleans up.
            string ps = @"
param([int]$RevitPid,[string]$Src,[string]$Dst,[string]$Marker,[string]$Version)
try { Wait-Process -Id $RevitPid -ErrorAction SilentlyContinue } catch {}
Start-Sleep -Seconds 2
$tries = 0
while ($tries -lt 10) {
  try {
    Copy-Item -Path (Join-Path $Src '*') -Destination $Dst -Recurse -Force -ErrorAction Stop
    Set-Content -Path $Marker -Value $Version -Encoding UTF8
    break
  } catch { Start-Sleep -Seconds 2; $tries++ }
}
Remove-Item -Path (Split-Path $Src -Parent) -Recurse -Force -ErrorAction SilentlyContinue
";
            File.WriteAllText(script, ps);

            string marker = Path.Combine(LocalRoot(year), "applied-version.txt");
            var psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File \""
                    + script + "\" -RevitPid " + pid
                    + " -Src \"" + payloadDir + "\""
                    + " -Dst \"" + liveDir + "\""
                    + " -Marker \"" + marker + "\""
                    + " -Version \"" + version + "\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
            };
            Process.Start(psi);
        }

        private static void TryDeleteDir(string dir)
        {
            try { Directory.Delete(dir, true); } catch { }
        }
    }
}
