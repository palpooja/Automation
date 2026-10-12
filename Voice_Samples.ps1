# URA Suite - voice samples for the "Job completed" announcement.
# Plays every English voice installed on this PC at three speeds and saves each sample as a .wav file
# in the "samples" folder next to this script, so you can listen again and choose.
# Uses only Windows' built-in speech (System.Speech) - nothing is installed or downloaded.

Add-Type -AssemblyName System.Speech
$synth  = New-Object System.Speech.Synthesis.SpeechSynthesizer
$synth.Volume = 100
$folder = Join-Path $PSScriptRoot "samples"
New-Item -ItemType Directory -Force -Path $folder | Out-Null

$voices = $synth.GetInstalledVoices() | Where-Object { $_.Enabled -and $_.VoiceInfo.Culture.Name -like "en*" } |
          ForEach-Object { $_.VoiceInfo }
if (-not $voices) {
    $voices = $synth.GetInstalledVoices() | Where-Object { $_.Enabled } | ForEach-Object { $_.VoiceInfo }
}
# Rate runs from -10 (slowest) to 10 (fastest); 0 is Windows' normal speed.
$speeds = @( @{ Name = "normal"; Rate = 0 }, @{ Name = "slow"; Rate = -2 }, @{ Name = "slower"; Rate = -4 } )

function New-Prompt([string[]]$parts) {
    $pb = New-Object System.Speech.Synthesis.PromptBuilder
    for ($i = 0; $i -lt $parts.Count; $i++) {
        if ($i -gt 0) { $pb.AppendBreak([TimeSpan]::FromMilliseconds(600)) }   # a calm pause between phrases
        $pb.AppendText($parts[$i])
    }
    return $pb
}

Write-Host ""
Write-Host "Voices found on this PC: $($voices.Count)" -ForegroundColor Cyan
$n = 0
foreach ($v in $voices) {
    foreach ($sp in $speeds) {
        $n++
        $label = "Sample {0}: {1} ({2}, {3}) - {4}" -f $n, $v.Name, $v.Culture.Name, $v.Gender, $sp.Name
        Write-Host $label -ForegroundColor Yellow
        $synth.SelectVoice($v.Name)
        $synth.Rate = $sp.Rate
        $between = New-Prompt @("Job completed.", "Starting another.")
        $last    = New-Prompt @("Job completed.")
        # 1) save to a file
        $safe = ($v.Name -replace '[^A-Za-z0-9]+', '_').Trim('_')
        $file = Join-Path $folder ("Sample_{0:D2}_{1}_{2}.wav" -f $n, $safe, $sp.Name)
        $synth.SetOutputToWaveFile($file)
        $synth.Speak($between)
        $pause = New-Object System.Speech.Synthesis.PromptBuilder
        $pause.AppendBreak([TimeSpan]::FromMilliseconds(1000))
        $synth.Speak($pause)
        $synth.Speak($last)
        # 2) play it aloud
        $synth.SetOutputToDefaultAudioDevice()
        $synth.Speak((New-Prompt @("Sample $n.")))
        Start-Sleep -Milliseconds 500
        $synth.Speak($between)
        Start-Sleep -Milliseconds 900
        $synth.Speak($last)
        Start-Sleep -Milliseconds 1500
    }
}
$synth.Dispose()
Write-Host ""
Write-Host "Done. The samples are saved in: $folder" -ForegroundColor Green
Write-Host "Tell Claude the sample number you like (for example 'Sample 5')." -ForegroundColor Green
Read-Host "Press Enter to close"
