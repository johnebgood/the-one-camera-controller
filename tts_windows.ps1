param (
    [string]$TextToSpeak = "Hello, this is the default text to speak."
)

Add-Type -AssemblyName System.Speech
$speechSynthesizer = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
$speechSynthesizer.Speak($TextToSpeak)
