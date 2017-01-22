Add-Type -AssemblyName System.Speech
$speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
$url = 'http://localhost:8080/'
$listener = new-object system.net.httplistener
$listener.prefixes.add($url)
$listener.start()

while($listener.islistening) {
    
    $context = $listener.getcontext()
    $request = $context.request
    $response = $context.response
    $urlInfo = $request.url
    if($urlInfo.localPath -eq "/tts") {
        $audio = New-Object System.IO.MemoryStream
        $speak.SetOutputToWaveStream($audio)
        $speakString = [uri]::UnescapeDataString($urlInfo.Query.Substring(1))
        $speak.speak($speakString)
        $audio.Position = 0
        $response.AddHeader("Content-disposition", "attachment; filename=" + $speakString.Substring(0,15) + ".wav")
        $audio.writeTo($response.OutputStream)
    } else {
        $response.statuscode = 404
    }
    $response.close()
}