"""Spoken announcement when a client's Chrome closes (user, 6 Oct 2026): voice = sample 1 or 4 of the
sample player (1st / 2nd English Windows voice, normal speed); 'Job completed. Starting another.' between
clients, 'Job completed.' after the last.  Windows' own speech only; runs in the background."""
from core import voice


def test_phrases():
    assert voice.phrase(ok=True, more=True) == ["Job completed.", "Starting another."]
    assert voice.phrase(ok=True, more=False) == ["Job completed."]
    assert voice.phrase(ok=False, more=True) == ["Job completed, please check the errors.", "Starting another."]
    assert voice.phrase(ok=False, more=False) == ["Job completed, please check the errors."]
    assert voice.phrase(stopped=True) == ["Job stopped."]


def test_samples_map_to_the_sample_player():
    assert voice.SAMPLES == {"1": (0, 0), "4": (1, 0)}       # (voice position, speed) as in Voice_Samples.ps1


def test_speaks_in_the_background_with_the_chosen_voice(monkeypatch):
    calls = []
    monkeypatch.setattr(voice, "IS_WINDOWS", True)
    monkeypatch.setattr(voice.subprocess, "Popen", lambda args, **kw: calls.append((args, kw)))
    assert voice.say(["Job completed.", "Starting another."], choice="4") is True
    args, kw = calls[0]
    assert args[:4] == ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass"]
    script = args[-1]
    assert "$i = 1" in script and "Rate = 0" in script and "Starting another." in script
    assert kw.get("creationflags", 0) != 0, "no console window flashes up"


def test_off_or_not_windows_says_nothing(monkeypatch):
    monkeypatch.setattr(voice.subprocess, "Popen", lambda *a, **k: (_ for _ in ()).throw(AssertionError("spoke")))
    monkeypatch.setattr(voice, "IS_WINDOWS", True)
    assert voice.say(["Job completed."], choice="off") is False
    monkeypatch.setattr(voice, "IS_WINDOWS", False)
    assert voice.say(["Job completed."], choice="1") is False


def test_text_cannot_break_out_of_the_script(monkeypatch):
    calls = []
    monkeypatch.setattr(voice, "IS_WINDOWS", True)
    monkeypatch.setattr(voice.subprocess, "Popen", lambda args, **kw: calls.append(args))
    voice.say(["it's done\"; Remove-Item C:\\x"], choice="1")
    assert "'it''s done\"; Remove-Item C:\\x'" in calls[0][-1], "quoted as one PowerShell string"
