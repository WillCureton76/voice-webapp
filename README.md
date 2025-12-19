# Voice Claude (VPS Edition)

Voice-first Claude assistant with streaming TTS and continuous conversation mode.

## Features

- **Voice input** via Web Speech API
- **Streaming TTS** via VPS Edge TTS endpoint (MSE streaming)
- **Continuous mode** - hands-free conversation loop
- **Daemon mode** - connect to VPS-hosted Claude Code
- **Hold-to-stop** gesture (400ms) for interrupting

## Deployment

Served from `/opt/voice-webapp/` on Will's VPS at `https://vps.willcureton.com/voice/`

## Files

- `index.html` - Light mode UI with mic bar
- `app.js` - All the logic (speech recognition, TTS, API calls)
