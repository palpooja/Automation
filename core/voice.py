"""core/voice.py - spoken announcement when a client's Chrome window has closed (user, 6 Oct 2026).

'Job completed. Starting another.' between clients, 'Job completed.' after the last one.
Voices = the user's choice from Voice_Samples.ps1: sample 1 (1st English voice, normal speed) or
sample 4 (2nd English voice, normal speed) - the same voice list and order as the sample player.
Windows' built-in speech (System.Speech through PowerShell): nothing to install, runs in the
background so the next client never waits, and never stops a run if speech is unavailable.
"""
import subprocess
import sys

IS_WINDOWS = sys.platform == "win32"
SAMPLES = {"1": (0, 0), "4": (1, 0)}          # sample number -> (position in the voice list, speed)
CHOICES = {"1": "Sample 1", "4": "Sample 4", "off": "Off"}
DEFAULT = "1"


def phrase(ok=True, more=False, stopped=False):
    if stopped:
        return ["Job stopped."]
    first = "Job completed." if ok else "Job completed, please check the errors."
    return [first, "Starting another."] if more else [first]


def attention_phrase():
    """Said when the program needs the person while it keeps working (10.14: a return slow to open)."""
    return ["Attention needed."]


def _quote(text):
    return "'" + str(text).replace("'", "''") + "'"           # one PowerShell literal string


def script(parts, choice):
    pos, rate = SAMPLES[choice]
    texts = ", ".join(_quote(p) for p in parts)
    return (
        "Add-Type -AssemblyName System.Speech; "
        "$s = New-Object System.Speech.Synthesis.SpeechSynthesizer; $s.Volume = 100; "
        "$v = @($s.GetInstalledVoices() | Where-Object { $_.Enabled -and $_.VoiceInfo.Culture.Name -like 'en*' }"
        " | ForEach-Object { $_.VoiceInfo }); "
        "if ($v.Count -eq 0) { $v = @($s.GetInstalledVoices() | Where-Object { $_.Enabled }"
        " | ForEach-Object { $_.VoiceInfo }) }; "
        f"$i = {pos}; if ($v.Count -gt 0) {{ if ($i -ge $v.Count) {{ $i = 0 }}; $s.SelectVoice($v[$i].Name) }}; "
        f"$s.Rate = {rate}; "
        "$pb = New-Object System.Speech.Synthesis.PromptBuilder; $first = $true; "
        f"foreach ($t in @({texts})) {{ if (-not $first) {{ $pb.AppendBreak([TimeSpan]::FromMilliseconds(600)) }}; "
        "$pb.AppendText($t); $first = $false }; "
        "$s.Speak($pb); $s.Dispose()")


def say(parts, choice=DEFAULT):
    """Speak in the background.  True if speech was started."""
    if not IS_WINDOWS or choice not in SAMPLES:
        return False
    try:
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
        subprocess.Popen(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command",
                          script(parts, choice)],
                         stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                         creationflags=flags)
        return True
    except Exception:                                       # speech is a convenience, never a failure
        return False
