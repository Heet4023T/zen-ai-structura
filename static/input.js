// ==========================================================================
// 1. FILE UPLOAD & DELETE LOGIC
// ==========================================================================
document.getElementById("imageInput").addEventListener("change", (e) => {
    const fileNameDisplay = document.getElementById("fileName");
    const deleteBtn = document.getElementById("deleteFileBtn");
    const restoredInput = document.getElementById("restoredFilename");
    if (e.target.files && e.target.files[0]) {
        fileNameDisplay.innerText = e.target.files[0].name;
        fileNameDisplay.style.color = "#e5e7eb"; 
        deleteBtn.style.display = "flex";
        restoredInput.value = ""; 
    }
});

function removeFile(event) {
    if (event) event.stopPropagation();
    document.getElementById("imageInput").value = ""; 
    const fileNameDisplay = document.getElementById("fileName");
    fileNameDisplay.innerText = "PNG, JPG supported";
    fileNameDisplay.style.color = "#94a3b8"; 
    document.getElementById("deleteFileBtn").style.display = "none";
    document.getElementById("restoredFilename").value = ""; 
}

// ==========================================================================
// 2. VOICE RECOGNITION LOGIC
// ==========================================================================
function startVoice() {
    const micStatus = document.getElementById("mic-status");
    const textarea = document.getElementById("prompt");
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRecognition) { alert("Voice recognition not supported."); return; }
    const recognition = new SpeechRecognition();
    recognition.lang = "en-IN"; recognition.interimResults = false; recognition.maxAlternatives = 1;
    micStatus.innerText = "Listening..."; recognition.start();
    recognition.onresult = (event) => { textarea.value += event.results[0][0].transcript + " "; micStatus.innerText = "Voice captured"; };
    recognition.onerror = () => { micStatus.innerText = "Voice error"; };
    recognition.onend = () => { setTimeout(() => { micStatus.innerText = "Tap to speak"; }, 1500); };
}

// ==========================================================================
// 3. MAIN PROCESSING LOGIC (SEND DATA)
// ==========================================================================
async function sendData() {
    const fileInput = document.getElementById("imageInput");
    const promptInput = document.getElementById("prompt");
    const restoredInput = document.getElementById("restoredFilename");
    const btn = document.querySelector(".primary-btn"); 
    const file = fileInput.files[0];
    const prompt = promptInput.value.trim();
    const restoredFile = restoredInput.value;

    if (!file && !prompt && !restoredFile) { alert("Please upload a bill or enter a prompt."); return; }

    const originalText = btn.innerText;
    btn.innerText = "Processing..."; btn.disabled = true;

    const formData = new FormData();
    if (file) formData.append("image", file);
    if (prompt) formData.append("prompt", prompt);
    if (restoredFile) formData.append("restored_filename", restoredFile);

    try {
        const response = await fetch("/process", { method: "POST", body: formData });
        if (response.status === 403) {
            const errData = await response.json(); alert(errData.error); window.location.href = "/login"; return;
        }
        if (!response.ok) { alert("Processing failed."); return;
        }

        const result = await response.json();
        if (result.status === "ok") {
            const trialCounter = document.getElementById('trial-counter');
            if (trialCounter && result.trials_left !== undefined) {
                trialCounter.innerText = result.trials_left + " free trials left";
                if (result.trials_left === 0) trialCounter.style.color = "#ef4444"; 
            }
            const link = document.createElement('a');
            link.href = `/download?filename=${encodeURIComponent(result.filename)}`;
            link.setAttribute('download', result.filename);
            document.body.appendChild(link); link.click(); document.body.removeChild(link);
            btn.innerText = "Download Again"; btn.style.background = "#3b82f6";
            btn.onclick = () => window.location.href = `/download?filename=${result.filename}`;
            btn.disabled = false;
        } else { alert("Processing error: " + result.error); btn.innerText = originalText; btn.disabled = false; }
    } catch (err) { console.error(err); alert("Server connection error."); btn.innerText = originalText; btn.disabled = false; }
}

// ==========================================================================
// 4. RESTORE FUNCTION (GLOBAL & DECODING)
// ==========================================================================
window.restore = function(encodedFilename, encodedPrompt) {
    const filename = decodeURIComponent(encodedFilename);
    const prompt = decodeURIComponent(encodedPrompt);

    if(prompt && prompt !== 'null') document.getElementById("prompt").value = prompt;
    
    if(filename && filename !== 'null') {
        const fileNameDisplay = document.getElementById("fileName");
        fileNameDisplay.innerText = "Restored: " + filename;
        fileNameDisplay.style.color = "#22c55e";
        document.getElementById("restoredFilename").value = filename;
        document.getElementById("deleteFileBtn").style.display = "flex";
    }
    
    // Close modals if open (requires checking if function exists)
    if (typeof closeModals === "function") closeModals();
};

document.addEventListener('DOMContentLoaded', () => {
    const data = sessionStorage.getItem("restoreData");
    if (data) {
        const parsed = JSON.parse(data);
        // Data is ALREADY decoded when saving in index.html, so we encode it back to pass to window.restore
        // OR we can just set fields directly here. Let's reuse window.restore for consistency.
        window.restore(encodeURIComponent(parsed.filename), encodeURIComponent(parsed.prompt));
        sessionStorage.removeItem("restoreData"); 
    }
});

function clearAll() { window.location.reload(); }