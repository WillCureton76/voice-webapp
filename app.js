// Voice Claude - Frontend Application

// Debug logging to Redis
const debugLog = [];
function logDebug(msg) {
  const entry = new Date().toISOString().substr(11,12) + ' ' + msg;
  debugLog.push(entry);
  console.log('[DBG]', entry);
  if (debugLog.length % 5 === 0) {
    fetch('https://vps.willcureton.com/run', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'X-Api-Key': VPS_API_KEY },
      body: JSON.stringify({ command: 'redis-cli SET tts_debug "' + debugLog.slice(-30).join('\n').replace(/"/g, '\"') + '"' })
    }).catch(() => {});
  }
}

const VPS_API_KEY = '88045c9ab91b6e313a24d71cc6fda505be45ac8e89706db45c86254516219a84';
// =======================
// DAEMON MODE (Talk to VPS Claude directly)
// =======================
const DAEMON_URL = 'https://vps.willcureton.com/claude-code/query';
const DAEMON_STREAM_URL = 'https://vps.willcureton.com/claude-code-v2/stream-v2';

// Session ID for daemon conversation continuity
let daemonSessionId = localStorage.getItem('daemonSessionId') || null;
let daemonAbortController = null;

function isDaemonMode() {
  const checkbox = document.getElementById('daemonMode');
  return checkbox ? checkbox.checked : false;
}

// Tool name to friendly display
const TOOL_DISPLAY_NAMES = {
  'Bash': '🔧 Running command...',
  'Read': '📖 Reading file...',
  'Write': '✏️ Writing file...',
  'Edit': '✏️ Editing file...',
  'Glob': '🔍 Searching files...',
  'Grep': '🔍 Searching code...',
  'WebFetch': '🌐 Fetching web...',
  'WebSearch': '🔍 Searching web...',
  'Task': '🤖 Spawning agent...'
};

async function sendMessageToDaemon(text, assistantMsg) {
  setStatus('🤖 Connecting to Daemon...', 'active');
  
  // Mark that we're waiting for a response
  responseComplete = false;

  let fullReply = '';

  try {
    const body = { message: text };
    if (daemonSessionId) {
      body.session_id = daemonSessionId;
    }

    daemonAbortController = new AbortController();
    const response = await fetch(DAEMON_STREAM_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'X-Api-Key': VPS_API_KEY },
      body: JSON.stringify(body),
      signal: daemonAbortController.signal
    });

    if (!response.ok) {
      throw new Error('Daemon error: ' + response.status);
    }

    const reader = response.body.getReader();
    const decoder = new TextDecoder();

    while (true) {
      const { done, value } = await reader.read();
      if (done || isStoppingRequested) break;

      const chunk = decoder.decode(value, { stream: true });
      const lines = chunk.split('\n');

      for (const line of lines) {
        if (line.startsWith('data: ')) {
          try {
            const data = JSON.parse(line.slice(6));

            // Session ID - save for continuity
            if (data.type === 'session' && data.session_id) {
              daemonSessionId = data.session_id;
              localStorage.setItem('daemonSessionId', daemonSessionId);
            }

            // Tool use - show what Claude is doing
            if (data.type === 'tool_use') {
              // Flush any pending TTS before tool runs
              flushTtsBuffer();
              const display = TOOL_DISPLAY_NAMES[data.tool] || `⚙️ Using ${data.tool}...`;
              setStatus(display, 'active');
            }

            // Text content
            if (data.type === 'text' && data.text) {
              fullReply += data.text;
              const textSpan = assistantMsg.querySelector('.message-text');
              if (textSpan) {
                textSpan.textContent = fullReply;
              } else {
                assistantMsg.textContent = fullReply;
              }
              scrollToBottom();
              setStatus('💬 Responding...', 'active');
              // Stream TTS by sentence as chunks arrive
              queueTextForSpeech(data.text);
            }

            // Result with session_id
            if (data.type === 'result') {
              if (data.session_id) {
                daemonSessionId = data.session_id;
                localStorage.setItem('daemonSessionId', daemonSessionId);
              }
              if (assistantMsg.classList.contains('streaming')) { assistantMsg.classList.remove('streaming'); if (!assistantMsg.querySelector('.replay-btn')) addReplayButton(assistantMsg); }
            }

            // Done
            if (data.type === 'done' || data.done) {
              if (assistantMsg.classList.contains('streaming')) { assistantMsg.classList.remove('streaming'); if (!assistantMsg.querySelector('.replay-btn')) addReplayButton(assistantMsg); }
            }
          } catch (e) {
            // Skip invalid JSON
          }
        }
      }
    }

    conversationHistory.push({ role: 'assistant', content: fullReply });
    saveHistory();

    // TTS already streamed by sentence above, just flush any remaining buffer
    if (autoSpeakCheckbox.checked) {
      flushTtsBuffer();
    } else {
      // No TTS, so response is complete now
      responseComplete = true;
    }
    apiResponseComplete = true;

    setStatus('Ready');
    return fullReply;
  } catch (error) {
    console.error('Daemon error:', error);
    responseComplete = true; // Reset on error
    const textSpan = assistantMsg.querySelector('.message-text');
    if (textSpan) {
      textSpan.textContent = 'Daemon error: ' + error.message;
    } else {
      assistantMsg.textContent = 'Daemon error: ' + error.message;
    }
    assistantMsg.classList.add('error');
    setStatus('Daemon error', 'error');
    throw error;
  }
}


// =======================
// State
// =======================
let isRecording = false;
let tapToSendBuffer = ''; // Accumulated transcript for tap-to-send mode
let recognition = null;
let currentUtterance = null;
let conversationHistory = [];
let abortController = null;

// TTS sentence queue (text) – sentences get added while processQueue runs
let ttsQueue = [];
let isSpeaking = false;
let ttsBuffer = ''; // Buffer for sentence-level TTS streaming

// Continuous mode: only trigger recording when response is FULLY complete
let responseComplete = true; // true when not waiting for API response
let apiResponseComplete = true; // true when API has finished streaming text
let userCancelledListening = false; // true when user manually tapped to stop listening

// DOM Elements
const chat = document.getElementById('chat');
const emptyState = document.getElementById('emptyState');
const status = document.getElementById('status');
const sendBtn = document.getElementById('sendBtn');
const textInput = document.getElementById('textInput');
const speakingIndicator = document.getElementById('speakingIndicator');
const settingsModal = document.getElementById('settingsModal');
const systemMessageInput = document.getElementById('systemMessage');
const autoSpeakCheckbox = document.getElementById('autoSpeak');
const modelSelect = document.getElementById('modelSelect');

// =======================
// MIC BAR STATE MANAGEMENT
// =======================
const micBar = document.getElementById('micBar');
const micIcon = document.getElementById('micIcon');
const micText = document.getElementById('micText');
const waveBars = document.getElementById('waveBars');
const spinnerEl = document.getElementById('spinner');
const interruptHint = document.getElementById('interruptHint');

const SPEAKER_ICON = '<path d="M3 9v6h4l5 5V4L7 9H3zm13.5 3c0-1.77-1.02-3.29-2.5-4.03v8.05c1.48-.73 2.5-2.25 2.5-4.02zM14 3.23v2.06c2.89.86 5 3.54 5 6.71s-2.11 5.85-5 6.71v2.06c4.01-.91 7-4.49 7-8.77s-2.99-7.86-7-8.77z"/>';
const MIC_ICON = '<path d="M12 14c1.66 0 3-1.34 3-3V5c0-1.66-1.34-3-3-3S9 3.34 9 5v6c0 1.66 1.34 3 3 3zm5.91-3c-.49 0-.9.36-.98.85C16.52 14.2 14.47 16 12 16s-4.52-1.8-4.93-4.15c-.08-.49-.49-.85-.98-.85-.61 0-1.09.54-1 1.14.49 3 2.89 5.35 5.91 5.78V20c0 .55.45 1 1 1s1-.45 1-1v-2.08c3.02-.43 5.42-2.78 5.91-5.78.1-.6-.39-1.14-1-1.14z"/>';

let isStoppingRequested = false;

function setMicBarState(state) {
  if (!micBar) return;
  if (state === 'speaking' && isStoppingRequested) return;
  
  micBar.classList.remove('idle', 'listening', 'processing', 'speaking', 'stopping');
  micBar.classList.add(state);
  
  if (micIcon) micIcon.style.display = 'none';
  if (waveBars) waveBars.style.display = 'none';
  if (spinnerEl) spinnerEl.style.display = 'none';
  if (interruptHint) interruptHint.style.display = 'none';
  
  const textInputRow = document.querySelector('.text-input-row');
  if (textInputRow) {
    const wasHidden = textInputRow.style.display === 'none';
    textInputRow.style.display = (state === 'idle') ? 'flex' : 'none';
    // Re-scroll when footer expands back to full size
    if (state === 'idle' && wasHidden) {
      setTimeout(() => scrollToBottom(), 50);
    }
  }
  
  switch(state) {
    case 'idle':
      if (micIcon) { micIcon.innerHTML = MIC_ICON; micIcon.style.display = 'block'; }
      if (micText) micText.textContent = 'Tap to speak';
      // Only trigger continuous mode when response is FULLY complete (not between sentences)
      if (isContinuousMode() && !isStoppingRequested && responseComplete && !userCancelledListening) {
        setTimeout(() => {
          if (isContinuousMode() && !isRecording && !isSpeaking && responseComplete && !userCancelledListening) {
            startRecording();
          }
        }, 500); // 500ms delay to avoid picking up tail audio
      }
      break;
    case 'listening':
      if (waveBars) waveBars.style.display = 'flex';
      if (micText) micText.textContent = 'Listening...';
      break;
    case 'processing':
      if (spinnerEl) spinnerEl.style.display = 'block';
      if (micText) micText.textContent = 'Thinking...';
      break;
    case 'speaking':
      if (micIcon) { micIcon.innerHTML = SPEAKER_ICON; micIcon.style.display = 'block'; }
      if (micText) micText.textContent = 'Speaking...';
      if (interruptHint) interruptHint.style.display = 'inline';
      break;
  }
}

function isContinuousMode() {
  const checkbox = document.getElementById('continuousMode');
  return checkbox ? checkbox.checked : false;
}

function toggleContinuousMode() {
  const checkbox = document.getElementById('continuousMode');
  const btn = document.getElementById('continuousModeBtn');
  if (checkbox && btn) {
    checkbox.checked = !checkbox.checked;
    btn.classList.toggle('active', checkbox.checked);
    localStorage.setItem('voiceClaude_continuous', checkbox.checked);
  }
}

function toggleTapToSend() {
  const checkbox = document.getElementById('tapToSendMode');
  const btn = document.getElementById('tapToSendBtn');
  if (checkbox && btn) {
    checkbox.checked = !checkbox.checked;
    btn.classList.toggle('active', checkbox.checked);
    localStorage.setItem('voiceClaude_tapToSend', checkbox.checked);
  }
}

function toggleTextInput() {
  const inputSection = document.getElementById('inputSection');
  const modeLabel = document.getElementById('modeLabel');
  if (inputSection) {
    inputSection.classList.toggle('expanded');
    // Swap label between type and voice
    if (modeLabel) {
      modeLabel.textContent = inputSection.classList.contains('expanded') ? 'voice' : 'type';
    }
    // Focus the text input when expanded
    if (inputSection.classList.contains('expanded')) {
      setTimeout(() => {
        const textInput = document.getElementById('textInput');
        if (textInput) textInput.focus();
      }, 100);
    }
  }
}

function initModeButtons() {
  // Restore continuous mode state
  const savedContinuous = localStorage.getItem('voiceClaude_continuous');
  const continuousCheckbox = document.getElementById('continuousMode');
  const continuousBtn = document.getElementById('continuousModeBtn');
  if (savedContinuous === 'true' && continuousCheckbox && continuousBtn) {
    continuousCheckbox.checked = true;
    continuousBtn.classList.add('active');
  }
  
  // Restore tap-to-send state
  const savedTapToSend = localStorage.getItem('voiceClaude_tapToSend');
  const tapToSendCheckbox = document.getElementById('tapToSendMode');
  const tapToSendBtn = document.getElementById('tapToSendBtn');
  if (savedTapToSend === 'true' && tapToSendCheckbox && tapToSendBtn) {
    tapToSendCheckbox.checked = true;
    tapToSendBtn.classList.add('active');
  }
}

// Initialize
document.addEventListener('DOMContentLoaded', () => {
  loadSettings();
  loadHistory();
  setupSpeechRecognition();
  initModeButtons();

  // Enter key to send
  textInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  });

  // Tap speaking indicator to stop TTS
  speakingIndicator.addEventListener('click', stopSpeaking);

  // Hold-to-stop gesture handling
  let micBarPressStart = 0;
  let holdTimer = null;
  let isHolding = false;
  const HOLD_THRESHOLD = 400;
  
  function setStoppingState() {
    isStoppingRequested = true;
    if (daemonAbortController) {
      try { daemonAbortController.abort(); } catch(_) {}
    }
    if (abortController) {
      try { abortController.abort(); } catch(_) {}
    }
    if (micBar) {
      micBar.classList.remove('idle', 'listening', 'processing', 'speaking');
      micBar.classList.add('stopping');
      if (micText) micText.textContent = 'Stopping...';
    }
  }
  
  function handlePressStart() {
    if (isSpeaking) {
      micBarPressStart = Date.now();
      isHolding = false;
      holdTimer = setTimeout(() => {
        isHolding = true;
        setStoppingState();
      }, HOLD_THRESHOLD);
    }
  }
  
  function handlePressEnd(e) {
    if (micBarPressStart > 0) {
      if (holdTimer) clearTimeout(holdTimer);
      const wasHolding = isHolding;
      micBarPressStart = 0;
      isHolding = false;
      
      if (isSpeaking || wasHolding) {
        stopSpeaking();
        if (!wasHolding) {
          setTimeout(() => startRecording(), 50);
        }
      }
    }
  }
  
  if (micBar) {
    micBar.addEventListener('touchstart', handlePressStart, { passive: true });
    micBar.addEventListener('mousedown', handlePressStart);
    micBar.addEventListener('touchend', function(e) { if (micBarPressStart > 0) e.preventDefault(); handlePressEnd(e); });
    micBar.addEventListener('mouseup', handlePressEnd);
    micBar.addEventListener('mouseleave', function() {
      if (holdTimer) clearTimeout(holdTimer);
      if (isHolding) { isHolding = false; if (isSpeaking) setMicBarState('speaking'); }
    });
  }
});

// =======================
// Speech Recognition Setup
// =======================
function setupSpeechRecognition() {
  if (!('webkitSpeechRecognition' in window) && !('SpeechRecognition' in window)) {
    setStatus('Speech recognition not supported', 'error');
    return;
  }

  const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  recognition = new SpeechRecognition();
  recognition.continuous = true;
  recognition.interimResults = true;
  recognition.lang = 'en-US';

  recognition.onstart = () => {
    isRecording = true;
    setMicBarState('listening');
    setStatus('Listening... tap mic when done', 'active');
  };

  recognition.onresult = (event) => {
    let transcript = '';
    for (let i = event.resultIndex; i < event.results.length; i++) {
      transcript += event.results[i][0].transcript;
    }
    
    // In tap-to-send mode, append to accumulated buffer
    const tapToSend = document.getElementById('tapToSendMode');
    if (tapToSend && tapToSend.checked && tapToSendBuffer) {
      textInput.value = tapToSendBuffer + ' ' + transcript;
    } else {
      textInput.value = transcript;
    }
  };

  recognition.onerror = (event) => {
    console.error('Speech recognition error:', event.error);
    stopRecording();
    if (event.error !== 'aborted') {
      setStatus('Error: ' + event.error, 'error');
    }
  };

  recognition.onend = () => {
    if (isRecording) {
      // Check if tap-to-send mode is enabled
      const tapToSend = document.getElementById('tapToSendMode');
      if (tapToSend && tapToSend.checked) {
        // Save current transcript before restarting
        tapToSendBuffer = textInput.value.trim();
        // Keep listening - restart recognition to accumulate more speech
        try {
          recognition.start();
        } catch (e) {
          // If start fails, fall back to normal behaviour
          tapToSendBuffer = ''; // Clear buffer
          stopRecording();
          if (textInput.value.trim()) {
            sendMessage();
          }
        }
      } else {
        // Normal behaviour: auto-send when recognition ends
        stopRecording();
        if (textInput.value.trim()) {
          sendMessage();
        }
      }
    }
  };
}

// Toggle Recording
function toggleRecording() {
  // If TTS is playing, stop it and start recording
  if (window.speechSynthesis?.speaking) {
    stopSpeaking();
  }

  if (isRecording) {
    stopRecording();
    // Send the message if we have text
    if (textInput.value.trim()) {
      sendMessage();
    }
  } else {
    startRecording();
  }
}

function startRecording() {
  if (!recognition) return;

  textInput.value = '';
  tapToSendBuffer = ''; // Clear accumulated buffer for fresh start
  try {
    recognition.start();
  } catch (e) {
    // Already started
  }
}

function stopRecording() {
  if (isRecording) {
    userCancelledListening = true; // User manually cancelled
  }
  isRecording = false;
  setMicBarState('idle');
  if (recognition) {
    try {
      recognition.stop();
    } catch (e) {
      // Already stopped
    }
  }
  setStatus('Ready');
}

// =======================
// Send Message
// =======================
async function sendMessage() {
  const text = textInput.value.trim();
  if (!text) return;

  // Clear tap-to-send buffer since we're sending
  tapToSendBuffer = '';

  // Mark that we're waiting for a response (prevents continuous mode triggering mid-response)
  responseComplete = false;

  // Cancel any ongoing request
  if (abortController) {
    abortController.abort();
  }

  // Stop any TTS and clear buffer
  stopSpeaking();
  ttsBuffer = '';
  apiResponseComplete = false; // Set AFTER stopSpeaking() to avoid race

  // Clear input
  textInput.value = '';

  // Hide empty state
  emptyState.style.display = 'none';

  // Add user message
  addMessage(text, 'user');
  conversationHistory.push({ role: 'user', content: text });
  saveHistory();

  // Create assistant message placeholder
  const assistantMsg = addMessage('', 'assistant', true);

  // Disable inputs
  setInputsEnabled(false);
  setStatus('Thinking...', 'active');

  try {
    // Check if daemon mode is enabled
    if (isDaemonMode()) {
      await sendMessageToDaemon(text, assistantMsg);
      return;
    }
    
    abortController = new AbortController();

    const response = await fetch('https://vps.willcureton.com/api/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'X-Api-Key': VPS_API_KEY },
      body: JSON.stringify({
        messages: conversationHistory,
        systemMessage: systemMessageInput.value || undefined,
        model: modelSelect.value
      }),
      signal: abortController.signal
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${await response.text()}`);
    }

    // Handle streaming response
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let fullResponse = '';
    let buffer = '';

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;

      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split('\n');
      buffer = lines.pop() || '';

      for (const line of lines) {
        if (line.startsWith('data: ')) {
          const data = line.slice(6);
          if (data === '[DONE]') continue;

          try {
            const parsed = JSON.parse(data);

            if (parsed.type === 'text') {
              fullResponse += parsed.content;
              const textSpan = assistantMsg.querySelector('.message-text');
              if (textSpan) {
                textSpan.textContent = fullResponse;
              } else {
                assistantMsg.textContent = fullResponse;
              }
              scrollToBottom();
              // Stream TTS by sentence as chunks arrive
              queueTextForSpeech(parsed.content);
            } else if (parsed.type === 'tool_call') {
              addMessage(`🔧 Calling: ${parsed.name}`, 'tool-call');
              setStatus(`Calling ${parsed.name}...`, 'active');
            } else if (parsed.type === 'tool_result') {
              addMessage(`✓ ${parsed.name}: ${parsed.result.substring(0, 100)}...`, 'tool-result');
            } else if (parsed.type === 'error') {
              throw new Error(parsed.message);
            }
          } catch (e) {
            if (e.message !== 'Unexpected end of JSON input') {
              console.error('Parse error:', e);
            }
          }
        }
      }
    }

    // Finalize
    if (assistantMsg.classList.contains('streaming')) { assistantMsg.classList.remove('streaming'); if (!assistantMsg.querySelector('.replay-btn')) addReplayButton(assistantMsg); }
    conversationHistory.push({ role: 'assistant', content: fullResponse });
    saveHistory();

    // Flush any remaining TTS buffer
    flushTtsBuffer();
    apiResponseComplete = true;
    
    // If auto-speak is off, response is complete now
    if (!autoSpeakCheckbox.checked) {
      responseComplete = true;
    }

    setStatus('Ready');
  } catch (error) {
    if (error.name === 'AbortError') {
      setStatus('Cancelled');
    } else {
      console.error('Error:', error);
      const textSpan = assistantMsg.querySelector('.message-text');
      if (textSpan) {
        textSpan.textContent = 'Error: ' + error.message;
      } else {
        assistantMsg.textContent = 'Error: ' + error.message;
      }
      assistantMsg.classList.add('error');
      setStatus('Error occurred', 'error');
    }
    responseComplete = true; // Reset on error/cancel
  } finally {
    setInputsEnabled(true);
    abortController = null;
  }
}

// Add message to chat
function addMessage(content, role, streaming = false) {
  const msg = document.createElement('div');
  msg.className = `message ${role}${streaming ? ' streaming' : ''}`;
  
  // Wrap text in span for assistant messages (allows button to have different styling)
  if (role === 'assistant') {
    const textSpan = document.createElement('span');
    textSpan.className = 'message-text';
    textSpan.textContent = content;
    msg.appendChild(textSpan);
  } else {
    msg.textContent = content;
  }
  
  // Add replay button for assistant messages (non-streaming)
  if (role === 'assistant' && !streaming) {
    addReplayButton(msg);
  }
  
  chat.appendChild(msg);
  scrollToBottom();
  return msg;
}

// Add replay button to a message
function addReplayButton(msgElement) {
  // Create actions wrapper
  const actions = document.createElement('div');
  actions.className = 'message-actions';
  
  // Play button
  const playBtn = document.createElement('button');
  playBtn.className = 'action-btn replay-btn';
  playBtn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="5 3 19 12 5 21 5 3"/></svg>';
  playBtn.title = 'Replay';
  playBtn.dataset.playing = 'false';
  playBtn.onclick = (e) => {
    e.stopPropagation();
    replayMessage(msgElement, playBtn);
  };
  
  // Flag button (placeholder)
  const flagBtn = document.createElement('button');
  flagBtn.className = 'action-btn flag-btn';
  flagBtn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 15s1-1 4-1 5 2 8 2 4-1 4-1V3s-1 1-4 1-5-2-8-2-4 1-4 1z"/><line x1="4" y1="22" x2="4" y2="15"/></svg>';
  flagBtn.title = 'Flag message (coming soon)';
  flagBtn.dataset.flagged = 'false';
  flagBtn.onclick = (e) => {
    e.stopPropagation();
    // Toggle flagged state visually
    if (flagBtn.dataset.flagged === 'true') {
      flagBtn.classList.remove('flagged');
      flagBtn.dataset.flagged = 'false';
    } else {
      flagBtn.classList.add('flagged');
      flagBtn.dataset.flagged = 'true';
    }
  };
  
  actions.appendChild(playBtn);
  actions.appendChild(flagBtn);
  msgElement.appendChild(actions);
}

// Replay a message's audio
let replayAudio = null;
let currentReplayBtn = null;

function stopReplay() {
  if (replayAudio) {
    replayAudio.pause();
    replayAudio = null;
  }
  if (currentReplayBtn) {
    currentReplayBtn.classList.remove('playing');
    currentReplayBtn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="5 3 19 12 5 21 5 3"/></svg>';
    currentReplayBtn.dataset.playing = 'false';
    currentReplayBtn = null;
  }
}

async function replayMessage(msgElement, btn) {
  // If this button is already playing, stop it
  if (btn.dataset.playing === 'true') {
    stopReplay();
    return;
  }
  
  // Stop any other playing audio
  stopReplay();
  
  // Get text content from message-text span or direct text nodes
  let text = '';
  const textSpan = msgElement.querySelector('.message-text');
  if (textSpan) {
    text = textSpan.textContent.trim();
  } else {
    for (const node of msgElement.childNodes) {
      if (node.nodeType === Node.TEXT_NODE) {
        text += node.textContent;
      }
    }
    text = text.trim();
  }
  
  if (!text) return;
  
  // Visual feedback - show stop icon in red
  btn.classList.add('playing');
  btn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="5" y="5" width="14" height="14"/></svg>';
  btn.dataset.playing = 'true';
  currentReplayBtn = btn;
  
  try {
    const response = await fetch(VPS_TTS_URL_FULL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-Api-Key': VPS_API_KEY
      },
      body: JSON.stringify({
        text: cleanTextForSpeech(text),
        voice: TTS_VOICE,
        rate: TTS_RATE,
        pitch: TTS_PITCH
      })
    });
    
    if (!response.ok) throw new Error('TTS failed: ' + response.status);
    
    const audioBlob = await response.blob();
    const audioUrl = URL.createObjectURL(audioBlob);
    replayAudio = new Audio(audioUrl);
    
    replayAudio.onended = () => {
      URL.revokeObjectURL(audioUrl);
      stopReplay();
    };
    
    replayAudio.onerror = (e) => {
      console.error('Audio error:', e);
      URL.revokeObjectURL(audioUrl);
      stopReplay();
    };
    
    await replayAudio.play();
  } catch (e) {
    console.error('Replay failed:', e);
    stopReplay();
  }
}

// =======================
// Text-to-Speech (REWRITTEN)
// MSE + true streaming + prefetch + sequencing with no gaps
// =======================

function cleanTextForSpeech(text) {
  return text
    .replace(/```[\s\S]*?```/g, 'code block')
    .replace(/`[^`]+`/g, '')
    .replace(/\*\*([^*]+)\*\*/g, '$1')
    .replace(/\*([^*]+)\*/g, '$1')
    .replace(/#{1,6}\s/g, '')
    .replace(/\[([^\]]+)\]\([^)]+\)/g, '$1')
    .replace(/[#*_~`]/g, '');
}

// Queue a chunk of text - extracts complete sentences and speaks them
function queueTextForSpeech(chunk) {
  if (!autoSpeakCheckbox.checked) return;

  ttsBuffer += chunk;

  const ABBREVIATIONS = /\b(Mr|Mrs|Ms|Dr|Jr|Sr|St|Prof|Inc|Ltd|Corp|vs|etc|i\.e|e\.g|U\.S|U\.K)\.\s*$/i;

  let sentences = [];
  let remaining = ttsBuffer;
  
  while (remaining.length > 0) {
    const match = remaining.match(/^(.*?[.!?:])(\s*)(.*)$/s);
    
    if (!match) {
      break;
    }
    
    const [, beforePunc, whitespace, afterPunc] = match;
    
    if (ABBREVIATIONS.test(beforePunc)) {
      if (whitespace && afterPunc) {
        const abbrevMatch = remaining.match(/^(.*?\b(?:Mr|Mrs|Ms|Dr|Jr|Sr|St|Prof|Inc|Ltd|Corp|vs|etc|i\.e|e\.g|U\.S|U\.K)\.\s*)(.*)$/is);
        if (abbrevMatch && abbrevMatch[2]) {
          remaining = abbrevMatch[1] + abbrevMatch[2];
          const realEnd = remaining.match(/^(.*?[.!?:])(\s+)(.*)$/s);
          if (realEnd) {
            sentences.push(realEnd[1]);
            remaining = realEnd[3];
            continue;
          }
        }
      }
      break;
    }
    
    if (whitespace) {
      sentences.push(beforePunc);
      remaining = afterPunc;
    } else if (afterPunc.length === 0) {
      if (remaining.length > 100) {
        sentences.push(beforePunc);
        remaining = '';
      } else {
        break;
      }
    } else if (/^[A-Z]/.test(afterPunc)) {
      sentences.push(beforePunc);
      remaining = afterPunc;
    } else {
      break;
    }
  }
  
  ttsBuffer = remaining;
  
  for (const sentence of sentences) {
    const cleaned = cleanTextForSpeech(sentence).trim();
    if (cleaned.length > 0) {
      ttsQueue.push(cleaned);
    }
  }
  
  processQueue();
}

function flushTtsBuffer() {
  if (ttsBuffer.trim().length > 2) {
    ttsQueue.push(cleanTextForSpeech(ttsBuffer.trim()));
    ttsBuffer = '';
    processQueue();
  }
}

// VPS TTS Configuration
const VPS_TTS_URL_STREAM = 'https://vps.willcureton.com/tts/stream'; // chunked MP3
const VPS_TTS_URL_FULL = 'https://vps.willcureton.com/tts'; // full MP3 fallback
const TTS_VOICE = 'en-GB-RyanNeural';
const TTS_RATE = '+12%';
const TTS_PITCH = '-1Hz';

// --- MSE Streaming Engine ---
const MSE_MIME_CANDIDATES = [
  'audio/mpeg',
  'audio/mp4; codecs="mp4a.40.2"'
];

let ttsSessionId = 0;
let ttsRunnerActive = false;

let mseAudioEl = null;
let mediaSource = null;
let sourceBuffer = null;

let streamingSentence = null;
let prefetchSentence = null;

let appendQueue = [];
let appendInProgress = false;

function getMseMimeType() {
  if (!('MediaSource' in window)) return null;
  for (const mime of MSE_MIME_CANDIDATES) {
    try {
      if (MediaSource.isTypeSupported(mime)) return mime;
    } catch (_) {}
  }
  return null;
}

function ensureMseAudio() {
  if (mseAudioEl && mediaSource && sourceBuffer) return true;

  const mime = getMseMimeType();
  if (!mime) return false;

  mseAudioEl = new Audio();
  mseAudioEl.preload = 'auto';
  mseAudioEl.autoplay = false;

  mediaSource = new MediaSource();
  const objectUrl = URL.createObjectURL(mediaSource);
  mseAudioEl.src = objectUrl;

  mediaSource.addEventListener('sourceopen', () => {
    if (!mediaSource || mediaSource.readyState !== 'open') return;
    try {
      sourceBuffer = mediaSource.addSourceBuffer(mime);
      sourceBuffer.mode = 'sequence';

      sourceBuffer.addEventListener('updateend', () => {
        appendInProgress = false;
        pumpAppendQueue();
      });

      sourceBuffer.addEventListener('error', (e) => {
        console.error('SourceBuffer error:', e);
      });

      pumpAppendQueue();
    } catch (e) {
      console.error('Failed to addSourceBuffer:', e);
    }
  });

  mediaSource.addEventListener('error', (e) => {
    console.error('MediaSource error:', e);
  });

  mseAudioEl.addEventListener('error', (e) => {
    console.error('MSE audio element error:', e);
  });

  return true;
}

function enqueueAppend(uint8) {
  appendQueue.push(uint8);
  pumpAppendQueue();
}

function pumpAppendQueue() {
  if (!sourceBuffer || !mediaSource || mediaSource.readyState !== 'open') return;
  if (appendInProgress) return;
  if (!appendQueue.length) return;
  if (sourceBuffer.updating) return;

  const chunk = appendQueue.shift();
  appendInProgress = true;
  try {
    sourceBuffer.appendBuffer(chunk);
  } catch (e) {
    appendInProgress = false;
    console.error('appendBuffer failed:', e);
  }
}

async function playMseIfNeeded() {
  if (!mseAudioEl) return;
  if (mseAudioEl.paused) {
    try {
      await mseAudioEl.play();
    } catch (e) {
      console.warn('MSE play() failed:', e);
    }
  }
}

function makeSentenceStreamJob(text, sessionId) {
  const ctrl = new AbortController();

  return {
    id: crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random(),
    text,
    sessionId,
    ctrl,
    reader: null,
    done: false,
    started: false,
    pumpPromise: null
  };
}

async function startSentenceStream(job) {
  const res = await fetch(VPS_TTS_URL_STREAM, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'X-Api-Key': VPS_API_KEY
    },
    body: JSON.stringify({
      text: job.text,
      voice: TTS_VOICE,
      rate: TTS_RATE,
      pitch: TTS_PITCH
    }),
    signal: job.ctrl.signal
  });

  if (!res.ok || !res.body) {
    throw new Error(`TTS stream failed: HTTP ${res.status}`);
  }

  job.reader = res.body.getReader();
  job.started = true;
  return job;
}

async function pumpSentenceToMse(job) {
  while (true) {
    if (job.sessionId !== ttsSessionId) throw new Error('Session changed');
    const { done, value } = await job.reader.read();
    if (done) {
      job.done = true;
      break;
    }
    if (value && value.byteLength) {
      enqueueAppend(value);
      playMseIfNeeded();
    }
  }
}

async function prefetchNextSentence(sessionId) {
  if (prefetchSentence || ttsQueue.length === 0) return;
  const nextText = ttsQueue[0];
  const job = makeSentenceStreamJob(nextText, sessionId);
  try {
    await startSentenceStream(job);
    if (sessionId !== ttsSessionId) {
      job.ctrl.abort();
      return;
    }
    prefetchSentence = job;
  } catch (e) {
    console.error('Prefetch stream error:', e);
    try { job.ctrl.abort(); } catch (_) {}
    prefetchSentence = null;
  }
}

function canEndStream() {
  if (!mediaSource || mediaSource.readyState !== 'open') return false;
  if (!apiResponseComplete) return false; // Don't end stream if API still sending
  if (ttsQueue.length !== 0) return false;
  if (streamingSentence && !streamingSentence.done) return false;
  if (prefetchSentence) return false;
  if (appendQueue.length !== 0) return false;
  if (sourceBuffer?.updating) return false;
  return true;
}

function tryFinalizeEndOfStream() {
  if (!mediaSource) return;
  if (!canEndStream()) return;

  try {
    mediaSource.endOfStream();
  } catch (e) {
    console.warn('endOfStream() failed:', e);
  }
}

async function processQueue() {
  logDebug('processQueue called: autoSpeak=' + autoSpeakCheckbox.checked + ' active=' + ttsRunnerActive + ' q=' + ttsQueue.length);
  if (!autoSpeakCheckbox.checked) return;
  if (ttsRunnerActive) { logDebug('SKIP: ttsRunnerActive'); return; }
  if (ttsQueue.length === 0) { logDebug('SKIP: queue empty'); return; }

  ttsRunnerActive = true;
  isSpeaking = true;
  setMicBarState('speaking');

  const mySession = ttsSessionId;

  try {
    if (!ensureMseAudio()) {
      while (ttsQueue.length > 0 && mySession === ttsSessionId) {
        const text = ttsQueue.shift();
        await new Promise((resolve) => {
          if (!('speechSynthesis' in window)) return resolve();
          const u = new SpeechSynthesisUtterance(text);
          u.rate = 1.15;
          u.pitch = 1.0;
          u.onend = resolve;
          u.onerror = resolve;
          window.speechSynthesis.speak(u);
        });
      }
      return;
    }

    logDebug('Waiting for MSE: ms=' + (mediaSource?.readyState) + ' sb=' + !!sourceBuffer);
    while (mySession === ttsSessionId && (!mediaSource || mediaSource.readyState !== 'open' || !sourceBuffer)) {
      await new Promise((r) => setTimeout(r, 10));
    }
    if (mySession !== ttsSessionId) { logDebug('Session changed, returning'); return; }
    logDebug('MSE ready, entering loop');

    while (mySession === ttsSessionId) {
      if (!streamingSentence) {
        if (ttsQueue.length === 0) { if (!apiResponseComplete) { await new Promise(r => setTimeout(r, 50)); continue; } logDebug('Loop exit: q empty, apiDone=' + apiResponseComplete); break; }

        const nextText = ttsQueue[0];

        if (prefetchSentence && prefetchSentence.text === nextText && prefetchSentence.sessionId === mySession) {
          streamingSentence = prefetchSentence;
          prefetchSentence = null;
          ttsQueue.shift();
        } else {
          const text = ttsQueue.shift();
          const job = makeSentenceStreamJob(text, mySession);
          streamingSentence = await startSentenceStream(job);
        }

        prefetchNextSentence(mySession);

        streamingSentence.pumpPromise = pumpSentenceToMse(streamingSentence)
          .catch((e) => {
            console.error('Sentence stream pump error:', e);
            streamingSentence.done = true;
          });
      }

      await streamingSentence.pumpPromise;
      streamingSentence = null;

      await new Promise((r) => setTimeout(r, 0));

      await prefetchNextSentence(mySession);
    }

    const finalizeSession = async () => {
      while (mySession === ttsSessionId && (appendQueue.length > 0 || sourceBuffer?.updating)) {
        await new Promise((r) => setTimeout(r, 10));
      }
      if (mySession !== ttsSessionId) return;
      tryFinalizeEndOfStream();
    };
    finalizeSession();

    if (mseAudioEl) {
      mseAudioEl.onended = () => {
        logDebug('ONENDED: session=' + mySession + '/' + ttsSessionId + ' q=' + ttsQueue.length + ' apiDone=' + apiResponseComplete);
        if (mySession !== ttsSessionId) return;
        isSpeaking = false;
        responseComplete = true; // Response AND TTS both done
        userCancelledListening = false; // 