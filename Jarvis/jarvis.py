import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser as wb
import os
import sys
import random
import pyautogui
import subprocess
import re
import shutil
import pyjokes
import threading
import queue
import time
import ctypes
import urllib.parse
import urllib.request
try:
    import vlc  # type: ignore
except Exception:
    vlc = None  # type: ignore
try:
    import yt_dlp  # type: ignore
except Exception:
    yt_dlp = None  # type: ignore

try:
    import tkinter as tk
    from tkinter.scrolledtext import ScrolledText
except Exception:
    tk = None
try:
    from dotenv import load_dotenv
except Exception:
    load_dotenv = None
try:
    # OpenAI
    from openai import OpenAI  # type: ignore
except Exception:
    OpenAI = None  # type: ignore
try:
    # Google Gemini
    import google.generativeai as genai  # type: ignore
except Exception:
    genai = None  # type: ignore
try:
    # Anthropic Claude
    import anthropic  # type: ignore
except Exception:
    anthropic = None  # type: ignore
try:
    import winreg  # Windows registry access
except Exception:
    winreg = None

# Hardcoded API defaults (env vars override when present)
DEFAULT_LLM_PROVIDER = os.environ.get('LLM_PROVIDER', '').strip() or 'openrouter'
DEFAULT_OPENROUTER_API_KEY = os.environ.get('OPENROUTER_API_KEY', '').strip() or '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'
DEFAULT_OPENROUTER_BASE = os.environ.get('OPENROUTER_BASE_URL', '').strip() or 'https://openrouter.ai/api/v1'
DEFAULT_OPENROUTER_MODEL = os.environ.get('OPENROUTER_MODEL', '').strip() or 'openai/gpt-4o'
try:
    import win32com.client  # for Windows startup shortcut and SAPI voice fallback
except Exception:
    win32com = None

engine = pyttsx3.init(driverName='sapi5')
voices = engine.getProperty('voices')
# Sensible defaults: prefer a female voice if available, else first voice
def _choose_default_voice_id():
    try:
        for v in voices:
            name = (getattr(v, 'name', '') or '').lower()
            if 'zira' in name or 'female' in name:
                return v.id
        return voices[0].id if voices else None
    except Exception:
        return None

if load_dotenv is not None:
    try:
        # Load .env and also fallback files if present
        load_dotenv()
        for alt in ('.jarvis.env', 'env.vars'):
            path = os.path.join(os.path.dirname(__file__), '..', alt)
            try:
                load_dotenv(os.path.abspath(path))
            except Exception:
                pass
    except Exception:
        pass

default_voice_id = _choose_default_voice_id()
if default_voice_id:
    engine.setProperty('voice', default_voice_id)
engine.setProperty('rate', 175)
engine.setProperty('volume', 1)


def set_voice_by_index(index: int) -> None:
    try:
        vs = engine.getProperty('voices')
        if 0 <= index < len(vs):
            engine.setProperty('voice', vs[index].id)
    except Exception:
        pass


def enable_system_voice() -> None:
    global USE_SAPI
    USE_SAPI = True


def enable_engine_voice() -> None:
    global USE_SAPI
    USE_SAPI = False


def _speak_sync(audio: str) -> None:
    global engine
    try:
        engine.say(audio)
        engine.runAndWait()
    except Exception:
        try:
            # Reinitialize engine once if needed
            engine = pyttsx3.init(driverName='sapi5')
            engine.say(audio)
            engine.runAndWait()
        except Exception:
            # Fallback to native SAPI voice if available
            try:
                if win32com is not None:
                    v = win32com.client.Dispatch('SAPI.SpVoice')
                    v.Rate = 1
                    v.Volume = 100
                    v.Speak(audio)
            except Exception:
                pass


class TTSWorker:
    def __init__(self):
        self._queue = queue.Queue()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def say(self, text: str) -> None:
        try:
            self._queue.put(str(text), block=False)
        except Exception:
            pass

    def _run(self) -> None:
        while True:
            try:
                text = self._queue.get()
                if text is None:
                    continue
                _speak_sync(text)
            except Exception:
                pass

    def stop(self) -> None:
        """Immediately stop current speech and clear any queued items."""
        try:
            try:
                engine.stop()
            except Exception:
                pass
            # Drain the queue
            while not self._queue.empty():
                try:
                    self._queue.get_nowait()
                except Exception:
                    break
        except Exception:
            pass


tts = TTSWorker()
USE_SAPI = True  # default to system voice for reliability
WAKE_ENABLED = True  # allow always-on wake word "jarvis"


def _sapi_speak(text: str) -> None:
    try:
        if win32com is not None:
            v = win32com.client.Dispatch('SAPI.SpVoice')
            v.Rate = 1
            v.Volume = 100
            v.Speak(text)
            return
    except Exception:
        pass
    _speak_sync(text)


def speak(audio) -> None:
    """Speak text using a dedicated worker; prefer native SAPI for maximum reliability."""
    message = str(audio)
    if USE_SAPI and win32com is not None:
        threading.Thread(target=_sapi_speak, args=(message,), daemon=True).start()
    else:
        tts.say(message)


def stop_speaking() -> None:
    """Stop the current TTS playback and clear queued speech."""
    try:
        tts.stop()
    except Exception:
        pass


def time_announce() -> None:
    """Tells the current time."""
    current_time = datetime.datetime.now().strftime("%I:%M:%S %p")
    speak("The current time is")
    speak(current_time)
    print("The current time is", current_time)
    try:
        if UI.instance is not None:
            UI.instance.log_assistant(f"The current time is {current_time}")
    except Exception:
        pass


def date_announce() -> None:
    """Tells the current date."""
    now = datetime.datetime.now()
    speak("The current date is")
    speak(f"{now.day} {now.strftime('%B')} {now.year}")
    print(f"The current date is {now.day}/{now.month}/{now.year}")
    try:
        if UI.instance is not None:
            UI.instance.log_assistant(f"The current date is {now.day} {now.strftime('%B')} {now.year}")
    except Exception:
        pass


def wishme() -> None:
    """Greets the user based on the time of day."""
    speak("Welcome back, sir!")
    print("Welcome back, sir!")

    hour = datetime.datetime.now().hour
    if 4 <= hour < 12:
        speak("Good morning!")
        print("Good morning!")
    elif 12 <= hour < 16:
        speak("Good afternoon!")
        print("Good afternoon!")
    elif 16 <= hour < 24:
        speak("Good evening!")
        print("Good evening!")
    else:
        speak("Good night, see you tomorrow.")

    assistant_name = load_name()
    speak(f"{assistant_name} at your service. Please tell me how may I assist you.")
    print(f"{assistant_name} at your service. Please tell me how may I assist you.")


def screenshot() -> None:
    """Takes a screenshot and saves it."""
    img = pyautogui.screenshot()
    img_path = os.path.expanduser("~\\Pictures\\screenshot.png")
    img.save(img_path)
    speak(f"Screenshot saved as {img_path}.")
    print(f"Screenshot saved as {img_path}.")


class UI:
    instance = None

    def __init__(self):
        if tk is None:
            UI.instance = None
            return
        self.root = tk.Tk()
        self.root.title("Jarvis Assistant")
        # Dark theme colors
        bg = "#121212"
        panel = "#1E1E1E"
        fg = "#E0E0E0"
        self.root.configure(bg=bg)
        self.text = ScrolledText(
            self.root,
            width=80,
            height=24,
            state='disabled',
            bg=panel,
            fg=fg,
            insertbackground=fg,
            borderwidth=0,
            highlightthickness=0
        )
        # Larger, readable font
        try:
            self.text.configure(font=("Segoe UI", 11))
        except Exception:
            pass
        self.text.pack(fill='both', expand=True)
        # Status bar
        self.status = tk.Label(self.root, text="Ready. Press N to talk.", anchor='w', bg=panel, fg="#B0BEC5")
        try:
            self.status.configure(font=("Segoe UI", 10))
        except Exception:
            pass
        self.status.pack(fill='x')
        # Tag styles
        self.text.tag_configure('user', foreground="#80CBC4")   # teal
        self.text.tag_configure('assistant', foreground="#C3E88D")  # green
        self.text.tag_configure('system', foreground="#9E9E9E")  # gray
        UI.instance = self

        # Bind push-to-talk (Press N)
        self.root.bind('<KeyPress-n>', self._on_ptt)
        # Add a button for mouse users
        self.button = tk.Button(self.root, text="Listen (N)", command=self._on_ptt_button, bg="#263238", fg="#ECEFF1", activebackground="#37474F", activeforeground="#FFFFFF", borderwidth=0)
        try:
            self.button.configure(font=("Segoe UI", 10, "bold"))
        except Exception:
            pass
        self.button.pack(pady=6)
        # Wake word toggle
        self.wake_var = tk.BooleanVar(value=True)
        self.wake_chk = tk.Checkbutton(self.root, text="Wake word: 'Jarvis'", variable=self.wake_var, command=self._toggle_wake, bg=bg, fg=fg, activebackground=bg, activeforeground=fg, selectcolor="#263238")
        try:
            self.wake_chk.configure(font=("Segoe UI", 10))
        except Exception:
            pass
        self.wake_chk.pack()
        # Audio test button
        self.test_btn = tk.Button(self.root, text="Audio Test", command=lambda: _speak_and_log("This is a voice test. If you can hear me, audio is working."), bg="#263238", fg="#ECEFF1", activebackground="#37474F", activeforeground="#FFFFFF", borderwidth=0)
        try:
            self.test_btn.configure(font=("Segoe UI", 10))
        except Exception:
            pass
        self.test_btn.pack(pady=2)
        # Stop Voice button
        self.stop_btn = tk.Button(self.root, text="Stop Voice (S)", command=stop_speaking, bg="#8E2430", fg="#FFFFFF", activebackground="#B33A44", activeforeground="#FFFFFF", borderwidth=0)
        try:
            self.stop_btn.configure(font=("Segoe UI", 10, "bold"))
        except Exception:
            pass
        self.stop_btn.pack(pady=2)
        self._listening_lock = threading.Lock()
        # Type-to-speak input
        bar = tk.Frame(self.root, bg=panel)
        bar.pack(fill='x', padx=6, pady=4)
        self.input_var = tk.StringVar()
        self.input = tk.Entry(bar, textvariable=self.input_var, bg="#263238", fg="#ECEFF1", insertbackground="#ECEFF1", relief='flat')
        try:
            self.input.configure(font=("Segoe UI", 10))
        except Exception:
            pass
        self.input.pack(side='left', fill='x', expand=True, padx=(0, 6))
        speak_btn = tk.Button(bar, text="Speak Text", command=lambda: _speak_and_log(self.input_var.get()), bg="#37474F", fg="#FFFFFF", borderwidth=0)
        try:
            speak_btn.configure(font=("Segoe UI", 10))
        except Exception:
            pass
        speak_btn.pack(side='right')
        # Bind Stop Voice (Press S)
        self.root.bind('<KeyPress-s>', lambda e: stop_speaking())

    def log(self, prefix: str, message: str) -> None:
        if tk is None:
            return
        self.text.configure(state='normal')
        tag = 'system'
        if prefix.lower().startswith('you'):
            tag = 'user'
        elif prefix.lower().startswith('assistant'):
            tag = 'assistant'
        self.text.insert('end', f"{prefix}: ", tag)
        self.text.insert('end', f"{message}\n")
        self.text.see('end')
        self.text.configure(state='disabled')

    def log_user(self, message: str) -> None:
        self.log("You", message)

    def log_assistant(self, message: str) -> None:
        self.log("Assistant", message)

    def run(self):
        if tk is not None:
            self.root.mainloop()

    def set_status(self, message: str) -> None:
        try:
            self.status.configure(text=message)
        except Exception:
            pass

    def _on_ptt(self, event):
        self._start_listen_once()

    def _on_ptt_button(self):
        self._start_listen_once()

    def _start_listen_once(self):
        if not self._listening_lock.acquire(blocking=False):
            return
        self.set_status("Listening… Speak now")
        threading.Thread(target=_listen_once_thread, daemon=True).start()

    def _toggle_wake(self):
        global WAKE_ENABLED
        WAKE_ENABLED = bool(self.wake_var.get())
        self.set_status("Wake word enabled" if WAKE_ENABLED else "Wake word disabled")


def takecommand() -> str | None:
    """Listen from the microphone and return recognized lowercased text or None."""
    recognizer = sr.Recognizer()
    recognizer.dynamic_energy_threshold = True
    recognizer.pause_threshold = 0.8
    recognizer.phrase_threshold = 0.2
    try:
        with sr.Microphone() as source:
            print("Listening… (calibrating background noise)")
            try:
                recognizer.adjust_for_ambient_noise(source, duration=0.6)
            except Exception:
                pass
            try:
                audio = recognizer.listen(source, timeout=8, phrase_time_limit=8)
            except sr.WaitTimeoutError:
                print("No speech detected in timeout window.")
                speak("I didn't hear anything. Please try again.")
                return None
    except OSError as e:
        print(f"Microphone error: {e}")
        speak("I couldn't access the microphone.")
        return None

    try:
        print("Recognizing…")
        text = recognizer.recognize_google(audio, language="en-US")
        print(text)
        try:
            if UI.instance is not None:
                UI.instance.log_user(text)
        except Exception:
            pass
        return text.lower()
    except sr.UnknownValueError:
        print("Unintelligible speech.")
        return None
    except sr.RequestError:
        print("Speech service unavailable.")
        speak("Speech recognition service is unavailable.")
        return None
    except Exception as e:
        print(f"Recognition error: {e}")
        return None


def play_music(song_name=None) -> None:
    """Play music by optional name; searches common user folders and CURRENT_DIR."""
    path = None
    if song_name:
        path = _find_audio_file(song_name)
    if path is None:
        # fallback: pick random from Music
        song_dir = os.path.expanduser("~\\Music")
        try:
            items = [f for f in os.listdir(song_dir) if _is_audio_file(f)]
        except Exception:
            items = []
        if items:
            path = os.path.join(song_dir, random.choice(items))
    if path and os.path.isfile(path):
        try:
            os.startfile(path)
            _speak_and_log(f"Playing {os.path.basename(path)}")
        except Exception:
            _speak_and_log("I found a track but couldn't play it.")
    else:
        _speak_and_log("I couldn't find that track.")


# ----- Media helpers -----
VK_MEDIA_NEXT_TRACK = 0xB0
VK_MEDIA_PREV_TRACK = 0xB1
VK_MEDIA_STOP = 0xB2
VK_MEDIA_PLAY_PAUSE = 0xB3


def _press_media_key(vk_code: int) -> None:
    try:
        ctypes.windll.user32.keybd_event(vk_code, 0, 0, 0)
        ctypes.windll.user32.keybd_event(vk_code, 0, 2, 0)  # KEYEVENTF_KEYUP = 2
    except Exception:
        pass


def pause_music() -> None:
    _press_media_key(VK_MEDIA_PLAY_PAUSE)
    _speak_and_log("Paused.")


def resume_music() -> None:
    _press_media_key(VK_MEDIA_PLAY_PAUSE)
    _speak_and_log("Resumed.")


def stop_music() -> None:
    _press_media_key(VK_MEDIA_STOP)
    _speak_and_log("Stopped.")


_AUDIO_EXTS = {".mp3", ".wav", ".flac", ".m4a", ".aac", ".ogg", ".wma"}


def _is_audio_file(filename: str) -> bool:
    try:
        return os.path.splitext(filename)[1].lower() in _AUDIO_EXTS
    except Exception:
        return False


def _find_audio_file(name: str) -> str | None:
    """Find an audio file by fuzzy name in CURRENT_DIR and common user folders."""
    q = _normalize_spoken_path_tokens(name).lower()
    roots = [
        CURRENT_DIR,
        os.path.expanduser("~\\Music"),
        os.path.expanduser("~\\Downloads"),
        os.path.expanduser("~\\Desktop"),
    ]
    # exact file name first
    for r in roots:
        try:
            cand = os.path.join(r, q)
            if os.path.isfile(cand):
                return cand
        except Exception:
            pass
    # fuzzy search by filename contains
    for r in roots:
        try:
            for base, _dirs, files in os.walk(r):
                for f in files:
                    if not _is_audio_file(f):
                        continue
                    if q in f.lower():
                        return os.path.join(base, f)
        except Exception:
            continue
    return None


def _play_from_web(song_query: str) -> bool:
    """Search YouTube for a song and open the first result with autoplay.

    Uses only the Python standard library. Falls back to normal search if needed.
    """
    try:
        q = urllib.parse.quote_plus(song_query.strip())
        search_url = f"https://www.youtube.com/results?search_query={q}"
        req = urllib.request.Request(search_url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=6) as resp:
            html = resp.read().decode("utf-8", errors="ignore")
        # Find a watch?v= video id
        import re as _re
        m = _re.search(r"watch\\?v=([A-Za-z0-9_-]{6,})", html)
        if not m:
            # Fallback: just open search results
            return _open_in_chrome_app(search_url)
        vid = m.group(1)
        # Use embed with autoplay and mute to comply with browser autoplay policies
        play_url = f"https://www.youtube.com/embed/{vid}?autoplay=1&mute=1&controls=1"
        return _open_in_chrome_app(play_url)
    except Exception:
        try:
            return _open_in_chrome_app(f"https://www.youtube.com/results?search_query={urllib.parse.quote_plus(song_query)}")
        except Exception:
            return False


def set_name() -> None:
    """Sets a new name for the assistant."""
    speak("What would you like to name me?")
    name = takecommand()
    if name:
        with open("assistant_name.txt", "w", encoding="utf-8") as file:
            file.write(name)
        speak(f"Alright, I will be called {name} from now on.")
    else:
        speak("Sorry, I couldn't catch that.")


def list_voices() -> None:
    """Lists available voices by index and name."""
    available = engine.getProperty('voices')
    lines = []
    for idx, v in enumerate(available):
        lines.append(f"{idx}: {getattr(v, 'name', 'Unknown')}")
    message = "; ".join(lines) if lines else "No voices found"
    print("Voices ->", message)
    speak("I have announced the list of available voices in the console.")


def set_voice_from_query(query: str) -> None:
    """Set voice based on user query (by index or gender keywords)."""
    available = engine.getProperty('voices')
    if not available:
        speak("No voices are available on this system.")
        return
    chosen = None
    # by index
    parts = query.split()
    for token in parts:
        if token.isdigit():
            i = int(token)
            if 0 <= i < len(available):
                chosen = available[i]
                break
    # by gender/name hints
    if chosen is None:
        q = query.lower()
        for v in available:
            name = (getattr(v, 'name', '') or '').lower()
            if ('female' in q and ('female' in name or 'zira' in name)) or \
               ('male' in q and ('male' in name or 'david' in name)) or \
               (name and any(tok.lower() in name for tok in parts)):
                chosen = v
                break
    if chosen is None:
        chosen = available[0]
    engine.setProperty('voice', chosen.id)
    speak("Voice updated.")


def load_name() -> str:
    """Loads the assistant's name from a file, or uses a default name."""
    try:
        with open("assistant_name.txt", "r", encoding="utf-8") as file:
            return file.read().strip()
    except FileNotFoundError:
        return "Jarvis"  # Default name


def search_wikipedia(query):
    """Searches Wikipedia and returns a summary."""
    try:
        speak("Searching Wikipedia...")
        result = wikipedia.summary(query, sentences=2)
        speak(result)
        print(result)
    except wikipedia.exceptions.DisambiguationError:
        speak("Multiple results found. Please be more specific.")
    except Exception:
        speak("I couldn't find anything on Wikipedia.")


def _current_model_label() -> str:
    p, _c, m = _select_llm_provider()
    if not p:
        return "none"
    return f"{p}:{m}"


def _speak_and_log(message: str) -> None:
    speak(message)
    try:
        if UI.instance is not None:
            UI.instance.log_assistant(message)
    except Exception:
        pass


CURRENT_DIR = os.getcwd()
PLAY_ON_WEB_BY_DEFAULT = False


class _OnlineAudioPlayer:
    _instance = None

    def __init__(self):
        self._lock = threading.RLock()
        self._playlist: list[tuple[str, str]] = []  # (url, title)
        self._index = -1
        self._vlc_instance = None
        self._player = None
        self._active = False

    @classmethod
    def instance(cls):
        if cls._instance is None:
            cls._instance = _OnlineAudioPlayer()
        return cls._instance

    def _ensure_vlc(self) -> bool:
        if vlc is None:
            _speak_and_log("Online player not installed. Please install requirements.")
            return False
        if self._vlc_instance is None:
            try:
                self._vlc_instance = vlc.Instance()
            except Exception:
                self._vlc_instance = None
        return self._vlc_instance is not None

    def _search_stream(self, query: str) -> tuple[str | None, str | None]:
        if yt_dlp is None:
            return None, None
        try:
            opts = {
                'format': 'bestaudio/best',
                'quiet': True,
                'no_warnings': True,
                'default_search': 'ytsearch1',
                'skip_download': True,
            }
            with yt_dlp.YoutubeDL(opts) as ydl:
                info = ydl.extract_info(query, download=False)
                if info is None:
                    return None, None
                if 'entries' in info:
                    info = info['entries'][0] if info['entries'] else None
                if not info:
                    return None, None
                stream_url = info.get('url')
                title = info.get('title') or query
                return stream_url, title
        except Exception:
            return None, None

    def _bind_events(self):
        try:
            if not self._player:
                return
            em = self._player.event_manager()
            em.event_detach_all()
            em.event_attach(vlc.EventType.MediaPlayerEndReached, lambda e: self.next())
        except Exception:
            pass

    def play_query(self, query: str) -> bool:
        if not self._ensure_vlc():
            return False
        url, title = self._search_stream(query)
        if not url:
            return False
        with self._lock:
            self._playlist.append((url, title or query))
            self._index = len(self._playlist) - 1
            try:
                self._player = self._vlc_instance.media_player_new()
                media = self._vlc_instance.media_new(url)
                self._player.set_media(media)
                self._player.audio_set_volume(85)
                self._bind_events()
                self._player.play()
                self._active = True
            except Exception:
                self._active = False
                return False
        _speak_and_log(f"Playing {title or 'music'} online")
        return True

    def pause(self) -> bool:
        with self._lock:
            if self._player and self._player.is_playing():
                try:
                    self._player.pause()
                    return True
                except Exception:
                    return False
        return False

    def resume(self) -> bool:
        with self._lock:
            if self._player and not self._player.is_playing():
                try:
                    self._player.play()
                    return True
                except Exception:
                    return False
        return False

    def stop(self) -> bool:
        with self._lock:
            if self._player:
                try:
                    self._player.stop()
                except Exception:
                    pass
            self._active = False
            return True

    def next(self) -> bool:
        with self._lock:
            if not self._playlist:
                return False
            # For now, replay the same or advance if more queued later
            self._index = min(self._index + 1, len(self._playlist) - 1)
            try:
                url, title = self._playlist[self._index]
                self._player.stop()
                media = self._vlc_instance.media_new(url)
                self._player.set_media(media)
                self._bind_events()
                self._player.play()
                return True
            except Exception:
                return False

    def is_active(self) -> bool:
        return bool(self._active and self._player)


def _play_online_background(song_query: str) -> bool:
    player = _OnlineAudioPlayer.instance()
    return player.play_query(song_query or "music")


def pause_online() -> bool:
    return _OnlineAudioPlayer.instance().pause()


def resume_online() -> bool:
    return _OnlineAudioPlayer.instance().resume()


def stop_online() -> bool:
    return _OnlineAudioPlayer.instance().stop()


def next_online() -> bool:
    return _OnlineAudioPlayer.instance().next()


def _open_in_chrome(url: str) -> None:
    try:
        chrome_path = r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
        if not os.path.exists(chrome_path):
            chrome_path = r"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"
        subprocess.Popen([chrome_path, url])
    except Exception:
        wb.open(url)


def _open_in_chrome_app(url: str) -> bool:
    """Open URL in Chrome app window (minimal UI) if Chrome is installed."""
    try:
        chrome_path = r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
        if not os.path.exists(chrome_path):
            chrome_path = r"C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"
        if os.path.exists(chrome_path):
            subprocess.Popen([chrome_path, f"--app={url}"])
            return True
    except Exception:
        pass
    try:
        wb.open(url)
        return True
    except Exception:
        return False


def _open_default_browser() -> None:
    try:
        # Try a blank page via webbrowser
        wb.open("about:blank")
    except Exception:
        try:
            # Fallback: open a safe URL
            os.startfile("https://www.google.com")
        except Exception:
            pass


def _open_discord() -> bool:
    """Attempt to open Discord using known Windows install layout."""
    try:
        local_app_data = os.environ.get("LOCALAPPDATA") or os.environ.get("LocalAppData")
        if local_app_data:
            updater = os.path.join(local_app_data, "Discord", "Update.exe")
            if os.path.isfile(updater):
                # Use Update.exe to start Discord by process name
                subprocess.Popen([updater, "--processStart", "Discord.exe"], shell=False)
                return True
        # Fallbacks: try protocol and direct names on PATH
        try:
            os.startfile("discord://")
            return True
        except Exception:
            pass
        try:
            subprocess.Popen(["Discord.exe"], shell=True)
            return True
        except Exception:
            pass
    except Exception:
        pass
    return False


def _find_app_executable(app_name: str) -> str | None:
    """Find an application executable by a friendly name.

    Search order:
    - Known aliases mapping
    - PATH (shutil.which)
    - Registry App Paths (HKCU/HKLM)
    - Start Menu shortcuts (*.lnk) under ProgramData and AppData
    - Common install directories under Program Files / LocalAppData
    """
    try:
        name = (app_name or "").strip().lower().strip('"')
        if not name:
            return None

        # Known aliases
        known: dict[str, list[str]] = {
            "discord": ["Discord.exe", "Discord"],
            "visual studio code": ["Code.exe", "code"],
            "vscode": ["Code.exe", "code"],
            "code": ["Code.exe", "code"],
            "chrome": ["chrome.exe"],
            "google chrome": ["chrome.exe"],
            "edge": ["msedge.exe"],
            "microsoft edge": ["msedge.exe"],
            "firefox": ["firefox.exe"],
            "spotify": ["Spotify.exe"],
            "steam": ["Steam.exe"],
            "whatsapp": ["WhatsApp.exe"],
            "notepad++": ["notepad++.exe"],
            "postman": ["Postman.exe"],
            "pycharm": ["pycharm64.exe", "pycharm.exe"],
            "android studio": ["studio64.exe", "studio.exe"],
            "intellij": ["idea64.exe", "idea.exe"],
            "git bash": ["git-bash.exe"],
            "vlc": ["vlc.exe"],
        }
        candidates: list[str] = []
        if name in known:
            candidates.extend(known[name])
        # Also try raw token
        base_token = name.replace(" ", "")
        candidates.extend([name, f"{name}.exe", base_token, f"{base_token}.exe"])  # heuristic

        # PATH lookup
        for c in candidates:
            from shutil import which
            p = which(c)
            if p:
                return p

        # Registry App Paths
        if winreg is not None:
            subkey = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths"
            for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
                try:
                    with winreg.OpenKey(hive, subkey) as root:
                        i = 0
                        while True:
                            try:
                                kname = winreg.EnumKey(root, i)
                            except OSError:
                                break
                            i += 1
                            lower = kname.lower()
                            if any(tok in lower for tok in (name, base_token)):
                                try:
                                    with winreg.OpenKey(root, kname) as k:
                                        val, _ = winreg.QueryValueEx(k, None)
                                        if val and os.path.isfile(val):
                                            return val
                                except Exception:
                                    pass
                except Exception:
                    pass

        # Start Menu shortcuts
        start_roots = [
            os.path.join(os.environ.get('PROGRAMDATA', ''), 'Microsoft', 'Windows', 'Start Menu', 'Programs'),
            os.path.join(os.environ.get('APPDATA', ''), 'Microsoft', 'Windows', 'Start Menu', 'Programs'),
        ]
        for root in start_roots:
            try:
                for base, _dirs, files in os.walk(root):
                    for f in files:
                        if not f.lower().endswith('.lnk'):
                            continue
                        lower = f.lower()
                        if name in lower or base_token in lower:
                            return os.path.join(base, f)
            except Exception:
                pass

        # Common install dirs
        common_dirs = [
            os.environ.get('ProgramFiles', ''),
            os.environ.get('ProgramFiles(x86)', ''),
            os.environ.get('LOCALAPPDATA', ''),
        ]
        for root in common_dirs:
            try:
                if not root or not os.path.isdir(root):
                    continue
                for base, _dirs, files in os.walk(root):
                    for f in files:
                        if f.lower().endswith('.exe'):
                            lower = f.lower()
                            if name in lower or base_token in lower:
                                return os.path.join(base, f)
            except Exception:
                pass
    except Exception:
        pass
    return None


def _natural_ack(text: str) -> None:
    _speak_and_log(text)


def _empathetic_response(query: str) -> str | None:
    """Return a kind, supportive response if the query sounds personal/sad."""
    q = (query or "").lower()
    triggers_any = (
        "how are you" in q or
        "bad day" in q or
        "sad" in q or
        "upset" in q or
        "lonely" in q or
        "stressed" in q or
        "anxious" in q or
        "depressed" in q or
        "not okay" in q or
        "feeling down" in q or
        "i feel bad" in q or
        "i'm tired" in q or
        "i am tired" in q or
        "burned out" in q or
        "heartbroken" in q or
        "i'm angry" in q or
        "i am angry" in q
    )
    if not triggers_any:
        return None
    options = [
        "I'm here with you. I'm sorry it's been rough—want to tell me more about it?",
        "That sounds heavy. You don't have to carry it alone; I'm listening if you want to share.",
        "It's okay to feel this way. Take a deep breath. Would a small break or a glass of water help?",
        "I'm on your side. One step at a time—we'll get through this together.",
        "Thanks for being honest with me. You matter more than you know. How can I support you right now?",
        "I believe in you. Even tough days end. Want me to put on calming music or open a journal?",
        "Sending you a little encouragement: you’ve done hard things before, and you can again.",
    ]
    # Special case for "how are you"
    if "how are you" in q:
        return random.choice([
            "I'm doing well, thank you. How are you feeling today?",
            "I'm here and ready to help. How are you holding up?",
            "Grateful to be with you. How's your day going?",
        ])
    return random.choice(options)


def _git_clone(url: str, dest: str = None) -> bool:
    try:
        if dest is None or not dest.strip():
            dest = os.path.join(os.getcwd(), os.path.basename(url.rstrip('/')).replace('.git', ''))
        subprocess.check_call(["git", "clone", url, dest], shell=False)
        return True
    except Exception:
        return False


def _git_init_commit_push(remote_url: str, message: str = "update") -> bool:
    try:
        subprocess.check_call(["git", "init"], shell=False)
        subprocess.check_call(["git", "add", "-A"], shell=False)
        subprocess.check_call(["git", "commit", "-m", message], shell=False)
        try:
            subprocess.check_call(["git", "remote", "remove", "origin"], shell=False)
        except Exception:
            pass
    except Exception:
        pass
    try:
        subprocess.check_call(["git", "remote", "add", "origin", remote_url], shell=False)
        subprocess.check_call(["git", "branch", "-M", "main"], shell=False)
        subprocess.check_call(["git", "push", "-u", "origin", "main"], shell=False)
        return True
    except Exception:
        return False


def _detect_project_type(path: str | None = None) -> str | None:
    """Detect project type in the given directory.
    Returns one of: 'npm', 'django', or None.
    """
    try:
        base = os.path.abspath(path or CURRENT_DIR)
        if os.path.isfile(os.path.join(base, "package.json")):
            return "npm"
        if os.path.isfile(os.path.join(base, "manage.py")) or any(f.endswith("settings.py") for f in os.listdir(base) if os.path.isdir(os.path.join(base, f))):
            return "django"
    except Exception:
        pass
    return None


def _run_project_command(intent: str, cwd: str | None = None) -> bool:
    """Run a high-level project command based on detected type.
    intent in {'install','start','migrate','makemigrations','collectstatic'}
    """
    target = os.path.abspath(cwd or CURRENT_DIR)
    ptype = _detect_project_type(target)
    try:
        if ptype == "npm":
            if intent == "install":
                return _run_command(["npm", "install"], cwd=target)
            if intent == "start":
                return _run_command(["npm", "start"], cwd=target)
            if intent in ("build", "run build"):
                return _run_command(["npm", "run", "build"], cwd=target)
        if ptype == "django":
            py = sys.executable or "python"
            if intent == "install":
                # Best-effort: install from requirements.txt if present
                req = os.path.join(target, "requirements.txt")
                if os.path.isfile(req):
                    return _run_command([py, "-m", "pip", "install", "-r", "requirements.txt"], cwd=target)
                return False
            if intent in ("migrate", "makemigrations", "collectstatic", "runserver", "start"):
                cmd = [py, "manage.py"]
                if intent == "start" or intent == "runserver":
                    cmd += ["runserver"]
                else:
                    cmd += [intent]
                return _run_command(cmd, cwd=target)
    except Exception:
        return False
    return False


def _create_folder(path: str) -> bool:
    try:
        os.makedirs(path, exist_ok=True)
        return True
    except Exception:
        return False


def _delete_folder(path: str) -> bool:
    try:
        if os.path.isdir(path):
            shutil.rmtree(path)
            return True
        if os.path.isfile(path):
            os.remove(path)
            return True
        return False
    except Exception:
        return False


def _rename_item(src: str, dst: str) -> bool:
    try:
        os.rename(src, dst)
        return True
    except Exception:
        return False


def _set_current_dir(path: str) -> bool:
    global CURRENT_DIR
    try:
        abs_path = os.path.abspath(path)
        if os.path.isdir(abs_path):
            CURRENT_DIR = abs_path
            return True
        return False
    except Exception:
        return False


def _find_first(root: str, name: str, want_file: bool = True):
    name_lower = name.lower()
    for base, dirs, files in os.walk(root):
        if want_file:
            for f in files:
                if f.lower() == name_lower or name_lower in f.lower():
                    return os.path.join(base, f)
        else:
            for d in dirs:
                if d.lower() == name_lower or name_lower in d.lower():
                    return os.path.join(base, d)
    return None


def _search_user_common_roots(name: str, want_file: bool = True):
    try:
        home = os.path.expanduser("~")
        roots = [
            CURRENT_DIR,
            home,
            os.path.join(home, "Desktop"),
            os.path.join(home, "Documents"),
            os.path.join(home, "Downloads"),
            os.path.join(home, "Pictures"),
            os.path.join(home, "Music"),
            os.path.join(home, "Videos"),
        ]
        for r in roots:
            if not r or not os.path.isdir(r):
                continue
            found = _find_first(r, name, want_file=want_file)
            if found:
                return found
    except Exception:
        pass
    return None


def _run_command(cmd: list[str], cwd: str = None) -> bool:
    try:
        subprocess.Popen(cmd, cwd=cwd or CURRENT_DIR, shell=False)
        return True
    except Exception:
        try:
            # Fallback to shell string
            subprocess.Popen(" ".join(cmd), cwd=cwd or CURRENT_DIR, shell=True)
            return True
        except Exception:
            return False


def _run_shell_string(command_str: str, cwd: str | None = None) -> bool:
    """Run an arbitrary shell command string in the given working directory.

    This uses `shell=True` to allow compound commands and builtins.
    """
    try:
        if not command_str or not command_str.strip():
            return False
        subprocess.Popen(command_str, cwd=cwd or CURRENT_DIR, shell=True)
        return True
    except Exception:
        return False


def _resolve_name_to_path(name: str, want_file: bool) -> str | None:
    """Resolve a spoken name to a file or folder path, searching CURRENT_DIR and user roots."""
    name = _normalize_spoken_path_tokens(name)
    # Try absolute or relative direct
    for cand in (name, os.path.join(CURRENT_DIR, name)):
        try:
            if want_file and os.path.isfile(cand):
                return os.path.abspath(cand)
            if not want_file and os.path.isdir(cand):
                return os.path.abspath(cand)
        except Exception:
            pass
    # Search trees
    found = _find_first(CURRENT_DIR, name, want_file=want_file) or _search_user_common_roots(name, want_file=want_file)
    return found


def _normalize_spoken_path_tokens(text: str) -> str:
    """Convert spoken tokens like 'underscore' -> '_', 'slash' -> '\\', 'space' -> ' '."""
    try:
        t = " " + (text or "") + " "
        # Replace common spoken separators
        replacements = [
            (" underscore ", "_"),
            (" dash ", "-"),
            (" hyphen ", "-"),
            (" dot ", "."),
            (" period ", "."),
            (" space ", " "),
            (" backslash ", "\\"),
            (" slash ", "\\"),  # favor Windows path
            (" forward slash ", "\\"),
        ]
        for k, v in replacements:
            t = t.replace(k, v)
        # Collapse extra spaces around separators
        t = t.strip().strip('"').strip()
        return t
    except Exception:
        return (text or "").strip().strip('"')


def _select_llm_provider():
    """Return a tuple (provider, client_or_model_callable, model_name) based on env vars.
    Provider is one of: 'openai', 'gemini', 'anthropic', or None if unavailable.
    """
    preferred = (os.environ.get('LLM_PROVIDER') or DEFAULT_LLM_PROVIDER).lower().strip()

    # OpenRouter (uses OpenAI SDK)
    if (preferred in ('openrouter', 'open-router', 'router')) and OpenAI is not None and ((os.environ.get('OPENROUTER_API_KEY') or DEFAULT_OPENROUTER_API_KEY)):
        try:
            base_url = os.environ.get('OPENROUTER_BASE_URL', DEFAULT_OPENROUTER_BASE)
            api_key = os.environ.get('OPENROUTER_API_KEY', DEFAULT_OPENROUTER_API_KEY)
            client = OpenAI(api_key=api_key, base_url=base_url)
            model = os.environ.get('OPENROUTER_MODEL', DEFAULT_OPENROUTER_MODEL)
            return 'openrouter', client, model
        except Exception:
            pass

    # OpenAI
    if (preferred in ('', 'openai')) and OpenAI is not None and os.environ.get('OPENAI_API_KEY'):
        try:
            client = OpenAI(api_key=os.environ.get('OPENAI_API_KEY'))
            model = os.environ.get('OPENAI_MODEL', 'gpt-4o-mini')
            return 'openai', client, model
        except Exception:
            pass

    # Gemini
    if (preferred in ('', 'gemini', 'google')) and genai is not None and os.environ.get('GOOGLE_API_KEY'):
        try:
            genai.configure(api_key=os.environ.get('GOOGLE_API_KEY'))
            model = os.environ.get('GEMINI_MODEL', 'gemini-1.5-flash')
            gm = genai.GenerativeModel(model)
            return 'gemini', gm, model
        except Exception:
            pass

    # Anthropic
    if (preferred in ('', 'anthropic', 'claude', 'xai')) and anthropic is not None and os.environ.get('ANTHROPIC_API_KEY'):
        try:
            client = anthropic.Anthropic(api_key=os.environ.get('ANTHROPIC_API_KEY'))
            model = os.environ.get('ANTHROPIC_MODEL', 'claude-3.5-sonnet')
            return 'anthropic', client, model
        except Exception:
            pass

    return None, None, None


def llm_generate_response(prompt: str, system_prompt: str = "You are a helpful desktop assistant.") -> str:
    provider, client, model = _select_llm_provider()
    if not provider:
        return "I can hear you. For open-ended questions, add API keys to enable smart answers."
    try:
        if provider in ('openai', 'openrouter'):
            extra_headers = None
            if provider == 'openrouter':
                referer = os.environ.get('OPENROUTER_SITE_URL', '')
                title = os.environ.get('OPENROUTER_SITE_NAME', '')
                extra_headers = {"HTTP-Referer": referer, "X-Title": title}
            resp = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": prompt},
                ],
                temperature=0.3,
                max_tokens=256,
                extra_headers=extra_headers,
            )
            return (resp.choices[0].message.content or "").strip()
        if provider == 'gemini':
            # google-generativeai
            full_prompt = f"{system_prompt}\n\nUser: {prompt}"
            resp = client.generate_content(full_prompt)
            return (getattr(resp, 'text', None) or "").strip()
        if provider == 'anthropic':
            resp = client.messages.create(
                model=model,
                max_tokens=256,
                temperature=0.3,
                system=system_prompt,
                messages=[{"role": "user", "content": prompt}],
            )
            # Anthropic returns a list of content blocks; concatenate text parts
            parts = []
            for b in getattr(resp, 'content', []) or []:
                t = getattr(b, 'text', None)
                if t:
                    parts.append(t)
            return ("\n".join(parts)).strip() or ""
    except Exception as e:
        return f"I can hear you, but I couldn't contact the model: {e}"
    return ""


def _handle_query(query: str) -> bool:
    global PLAY_ON_WEB_BY_DEFAULT
    # Wake-word handling: "jarvis" acknowledges and optionally strips the name
    if "jarvis" in query:
        cleaned = query
        # Normalize common wake phrases
        for wake in ("hey jarvis", "ok jarvis", "okay jarvis", "jarvis"):
            if cleaned.startswith(wake):
                cleaned = cleaned[len(wake):].strip(" ,.")
                break
        if not cleaned:
            _natural_ack("Yes, I'm listening.")
            return True
        _natural_ack("Sure")
        query = cleaned

    """Handle a single recognized query. Return True to continue loop, False to exit."""
    if not query:
        return True

    if "time" in query:
        time_announce()
    elif "date" in query:
        date_announce()
    elif "wikipedia" in query:
        q = query.replace("wikipedia", "").strip()
        if q:
            search_wikipedia(q)
        else:
            speak("What should I search on Wikipedia?")
    elif "play music" in query or query.startswith("play song ") or query.startswith("play track ") or "play online" in query or "play on web" in query or "play from web" in query:
        song_name = query
        for k in ("play music", "play song", "play track"):
            if k in query:
                song_name = query.split(k, 1)[1]
        song_name = song_name.strip()
        wants_online = PLAY_ON_WEB_BY_DEFAULT or any(x in query for x in ("play online", "play on web", "play from web", "online"))
        if wants_online:
            if _play_online_background(song_name or "music"):
                pass
            elif _play_from_web(song_name or "music"):
                _speak_and_log("Playing on the web")
            else:
                _speak_and_log("I couldn't play that online.")
        else:
            play_music(song_name)
    elif "always play online" in query or "play online by default" in query:
        PLAY_ON_WEB_BY_DEFAULT = True
        _speak_and_log("Okay, I will play music online by default.")
    elif "play locally" in query or "don't play online" in query or "do not play online" in query:
        PLAY_ON_WEB_BY_DEFAULT = False
        _speak_and_log("Okay, I will play music locally by default.")
    elif "pause music" in query or "pause song" in query or query.strip() == "pause":
        if not pause_online():
            pause_music()
        else:
            _speak_and_log("Paused.")
    elif "resume music" in query or "continue music" in query or query.strip() == "resume":
        if not resume_online():
            resume_music()
        else:
            _speak_and_log("Resumed.")
    elif "stop music" in query or "stop song" in query or query.strip() == "stop":
        if not stop_online():
            stop_music()
        else:
            _speak_and_log("Stopped.")
    elif "next song" in query or query.strip() == "next":
        if not next_online():
            _press_media_key(VK_MEDIA_NEXT_TRACK)
        _speak_and_log("Next.")
    elif "open youtube" in query:
        wb.open("youtube.com")
    elif "open google" in query:
        wb.open("https://www.google.com")
        _speak_and_log("Opening Google")
    elif ((query.startswith("open ") and not (query.startswith("open folder ") or query.startswith("open folder named ") or query.startswith("open directory ") or query.startswith("open directory named "))) or
          query.startswith("launch ") or
          query in ("open browser", "open the browser", "open web browser")):
        # Open application by name or path
        if query in ("open browser", "open the browser", "open web browser"):
            _open_default_browser()
            _speak_and_log("Opening your browser")
            return True
        target = query.replace("open ", "", 1).replace("launch ", "", 1).strip().strip('"')
        target = _normalize_spoken_path_tokens(target)
        opened = False
        # Try common system apps
        common = {
            "notepad": "notepad.exe",
            "calculator": "calc.exe",
            "paint": "mspaint.exe",
            "explorer": "explorer.exe",
            "cmd": "cmd.exe",
            "powershell": "powershell.exe",
            "chrome": r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
            "edge": r"C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
            "vlc": r"C:\\Program Files\\VideoLAN\\VLC\\vlc.exe",
            "discord": "discord",
        }
        try:
            tl = target.lower()
            if tl == "discord":
                if _open_discord():
                    _speak_and_log("Opening Discord")
                    opened = True
            # If target is a folder path, open it in Explorer
            if not opened and os.path.isdir(target):
                dest = os.path.abspath(target)
                try:
                    os.startfile(dest)
                except Exception:
                    pass
                try:
                    _set_current_dir(dest)
                except Exception:
                    pass
                _speak_and_log(f"Opening folder {target}")
                opened = True
            if not opened:
                exe = common.get(tl) or _find_app_executable(target) or target
                # If it's a .lnk, let shell handle it; otherwise, try to spawn
                if isinstance(exe, str) and exe.lower().endswith('.lnk'):
                    os.startfile(exe)
                else:
                    subprocess.Popen(exe if (isinstance(exe, str) and (exe.endswith('.exe') or ' ' in exe)) else [exe], shell=True)
                _speak_and_log(f"Opening {target}")
                opened = True
        except Exception:
            pass
        if not opened:
            # Try to resolve as a folder name by search
            found_dir = _find_first(CURRENT_DIR, target, want_file=False) or _search_user_common_roots(target, want_file=False)
            if found_dir and os.path.isdir(found_dir):
                try:
                    os.startfile(found_dir)
                except Exception:
                    pass
                try:
                    _set_current_dir(found_dir)
                except Exception:
                    pass
                _speak_and_log(f"Opening folder {os.path.basename(found_dir)}")
            else:
                _speak_and_log(f"I couldn't open {target}.")
    elif query.startswith("create folder ") or query.startswith("make folder "):
        name = query.split("folder", 1)[1].strip().strip('"')
        path = os.path.abspath(name)
        if _create_folder(path):
            _speak_and_log(f"Created folder {name}")
        else:
            _speak_and_log(f"I couldn't create folder {name}")
    elif query.startswith("delete folder ") or query.startswith("remove folder ") or query.startswith("delete file "):
        name = query.split(" ", 2)[2].strip().strip('"')
        path = os.path.abspath(name)
        if _delete_folder(path):
            _speak_and_log(f"Deleted {name}")
        else:
            _speak_and_log(f"I couldn't delete {name}")
    elif query.startswith("rename "):
        # rename <old> to <new>
        m = re.search(r"rename\s+\"?([^\"]+)\"?\s+to\s+\"?([^\"]+)\"?", query)
        if m:
            old, new = m.group(1).strip(), m.group(2).strip()
            if _rename_item(old, new):
                _speak_and_log(f"Renamed {old} to {new}")
            else:
                _speak_and_log(f"I couldn't rename {old}")
        else:
            _speak_and_log("Please say rename <old> to <new>.")
    elif query.startswith("git clone "):
        url = query.split("git clone", 1)[1].strip()
        if _git_clone(url):
            _speak_and_log("Repository cloned.")
        else:
            _speak_and_log("I couldn't clone that repository.")
    elif query.startswith("push to github") or query.startswith("git push") or query.startswith("commit and push"):
        # Try to extract remote URL and commit message; if missing, prompt via UI or console
        m = re.search(r"(https?://\S+\.git)", query)
        remote = m.group(1) if m else None
        msg = None
        cm = re.search(r"message\s+\"([^\"]+)\"", query)
        if cm:
            msg = cm.group(1)
        if remote is None:
            # Prompt user
            if tk is not None and UI.instance is not None:
                try:
                    # Simple blocking prompt using a tiny dialog
                    import tkinter.simpledialog as sd
                    remote = sd.askstring("Git Remote", "Enter remote URL (ends with .git):")
                except Exception:
                    remote = None
            if remote is None:
                try:
                    remote = input("Enter remote URL (ends with .git): ").strip()
                except Exception:
                    remote = None
        if not msg:
            msg = "update"
        if not remote:
            _speak_and_log("I need a remote URL ending with .git to push.")
        else:
            if _git_init_commit_push(remote, message=msg):
                _speak_and_log("Changes committed and pushed to GitHub.")
            else:
                _speak_and_log("I couldn't push to GitHub.")
    elif query.startswith("type "):
        # Type text at cursor position
        text = query.replace("type ", "", 1)
        try:
            pyautogui.typewrite(text)
            _speak_and_log("Typed your text.")
        except Exception:
            _speak_and_log("I couldn't type that.")
    elif query.startswith("search "):
        # Search the web (default browser)
        q = query.replace("search ", "", 1)
        wb.open(f"https://www.google.com/search?q={q}")
        _speak_and_log(f"Searching for {q}")
    elif "amazon" in query and ("search" in query or "find" in query):
        # Simple Amazon search helper
        import urllib.parse as _u
        # Extract following words after 'search' or 'find'
        kw = query
        for k in ("search", "find", "for"):
            if k in query:
                kw = query.split(k, 1)[1]
                break
        kw = _u.quote_plus(kw.strip())
        wb.open(f"https://www.amazon.in/s?k={kw}")
        _speak_and_log("Opening Amazon search")
    elif ("youtube" in query and ("search" in query or "find" in query)) or ("in chrome" in query and "youtube" in query):
        # YouTube search helper, optionally target Chrome
        import urllib.parse as _u
        kw = query
        for k in ("search for", "search", "find"):
            if k in query:
                kw = query.split(k, 1)[1]
                break
        kw = _u.quote_plus(kw.strip())
        url = f"https://www.youtube.com/results?search_query={kw}"
        if "in chrome" in query or "in google chrome" in query:
            _open_in_chrome(url)
        else:
            wb.open(url)
        _speak_and_log("Opening YouTube search")
    elif ("in chrome" in query or "in google chrome" in query) and ("open" in query or "go to" in query):
        # Generic "in chrome open <site>" handler
        # Extract a URL-ish token after 'open'/'go to'
        m = re.search(r"(?:open|go to)\s+([\w\.-]+\.[a-z]{2,})(?:\s|$)", query)
        if m:
            host = m.group(1)
            if not host.startswith("http"):
                url = f"https://{host}"
            else:
                url = host
            _open_in_chrome(url)
            _speak_and_log(f"Opening {host} in Chrome")
        else:
            _speak_and_log("Please specify a website to open in Chrome")
    elif query.startswith("go to folder ") or query.startswith("cd to ") or query.startswith("open folder ") or query.startswith("open folder named "):
        # Accept variants like: open folder X, open folder named X
        if query.startswith("open folder named "):
            name = query.replace("open folder named ", "", 1).strip()
        else:
            name = query.split(" ", 2)[2].strip()
        name = _normalize_spoken_path_tokens(name)
        # try absolute, then relative to CURRENT_DIR, then search from CURRENT_DIR
        candidates = [name, os.path.join(CURRENT_DIR, name)]
        dest = None
        for c in candidates:
            if os.path.isdir(c):
                dest = c
                break
        if dest is None:
            found = _find_first(CURRENT_DIR, name, want_file=False)
            if found:
                dest = found
        if dest is None:
            # Search common user roots like Desktop, Documents, Downloads
            found = _search_user_common_roots(name, want_file=False)
            if found:
                dest = found
        if dest and _set_current_dir(dest):
            _speak_and_log(f"Moved to {dest}")
            try:
                # Also open in Explorer for visual confirmation
                os.startfile(dest)
            except Exception:
                pass
        else:
            _speak_and_log(f"I couldn't find folder {name}")
    elif query.startswith("go to ") and " and run " in query:
        # Example: go to C:\\Projects\\myapp and run npm start
        try:
            parts = query.replace("go to ", "", 1).split(" and run ", 1)
            folder = parts[0].strip().strip('"')
            cmd = parts[1].strip()
            # Resolve folder
            resolved = folder
            if not os.path.isabs(resolved):
                # try relative to CURRENT_DIR
                candidate = os.path.join(CURRENT_DIR, resolved)
                if os.path.isdir(candidate):
                    resolved = candidate
            if not os.path.isdir(resolved):
                found = _find_first(CURRENT_DIR, folder, want_file=False)
                if not found:
                    found = _search_user_common_roots(folder, want_file=False)
                if found:
                    resolved = found
            if os.path.isdir(resolved):
                if _run_shell_string(cmd, cwd=resolved):
                    _speak_and_log(f"Running {cmd} in {resolved}")
                else:
                    _speak_and_log("I couldn't run that command.")
            else:
                _speak_and_log(f"I couldn't find folder {folder}")
        except Exception:
            _speak_and_log("Please say: go to <folder> and run <command>.")
    elif query.startswith("in that open ") or query.startswith("in that go to "):
        # Example: in that open logs, in that go to src
        name = query.split(" ", 2)[2].strip().strip('"')
        dest = _resolve_name_to_path(name, want_file=False)
        if dest and _set_current_dir(dest):
            try:
                os.startfile(dest)
            except Exception:
                pass
            _speak_and_log(f"Opened folder {os.path.basename(dest)}")
        else:
            _speak_and_log(f"I couldn't find folder {name}")
    elif query.startswith("in that run ") or query.startswith("in that execute "):
        # Example: in that run npm start, in that run python app.py
        cmd = query.split(" ", 3)[3].strip()
        if _run_shell_string(cmd, cwd=CURRENT_DIR):
            _speak_and_log(f"Running {cmd}")
        else:
            _speak_and_log("I couldn't run that command.")
    elif query.startswith("in that open file ") or query.startswith("in that find file "):
        # Example: in that open file app.py
        name = query.split(" ", 4)[4].strip().strip('"')
        fp = _resolve_name_to_path(name, want_file=True)
        if fp and os.path.isfile(fp):
            try:
                os.startfile(fp)
                _speak_and_log(f"Opened {os.path.basename(fp)}")
            except Exception:
                _speak_and_log("I found it but could not open the file.")
        else:
            _speak_and_log(f"I couldn't find {name}")
    elif query.startswith("go to ") and " and run " in query:
        # Example: go to C:\Projects\myapp and run npm start
        try:
            parts = query.replace("go to ", "", 1).split(" and run ", 1)
            folder = parts[0].strip().strip('"')
            cmd = parts[1].strip()
            # Resolve folder
            resolved = folder
            if not os.path.isabs(resolved):
                # try relative to CURRENT_DIR
                candidate = os.path.join(CURRENT_DIR, resolved)
                if os.path.isdir(candidate):
                    resolved = candidate
            if not os.path.isdir(resolved):
                found = _find_first(CURRENT_DIR, folder, want_file=False)
                if found:
                    resolved = found
            if os.path.isdir(resolved):
                if _run_shell_string(cmd, cwd=resolved):
                    _speak_and_log(f"Running {cmd} in {resolved}")
                else:
                    _speak_and_log("I couldn't run that command.")
            else:
                _speak_and_log(f"I couldn't find folder {folder}")
        except Exception:
            _speak_and_log("Please say: go to <folder> and run <command>.")
    elif query.startswith("run ") or query.startswith("execute "):
        # Run arbitrary command in CURRENT_DIR
        cmd = query.split(" ", 1)[1].strip()
        if _run_shell_string(cmd, cwd=CURRENT_DIR):
            _speak_and_log(f"Running {cmd}")
        else:
            _speak_and_log("I couldn't run that command.")
    elif query.startswith("find file ") or query.startswith("open file ") or query.startswith("open file named "):
        # Support: open file X, open file named X, find file X
        name = query
        if query.startswith("open file named "):
            name = query.replace("open file named ", "", 1)
        else:
            name = query.split(" ", 2)[2]
        name = _normalize_spoken_path_tokens(name)
        # First: try exact/partial match in CURRENT_DIR
        found = _find_first(CURRENT_DIR, name, want_file=True)
        # If user says 'in my pc', search common user folders too
        if (not found) and (" in my pc" in query or " in my pc." in query):
            found = _search_user_common_roots(name, want_file=True)
        if found:
            try:
                os.startfile(found)
                _speak_and_log(f"Opened {os.path.basename(found)}")
            except Exception:
                _speak_and_log("I found it but could not open the file.")
        else:
            _speak_and_log(f"I couldn't find {name}")
    elif query.startswith("npm install") or query.startswith("install npm"):
        if _run_project_command("install", cwd=CURRENT_DIR) or _run_command(["npm", "install"], cwd=CURRENT_DIR):
            _speak_and_log("Running install in this project")
        else:
            _speak_and_log("I couldn't run install here")
    elif query.startswith("npm start") or "run the app" in query or "start the app" in query:
        if _run_project_command("start", cwd=CURRENT_DIR) or _run_command(["npm", "start"], cwd=CURRENT_DIR):
            _speak_and_log("Starting the app")
        else:
            _speak_and_log("I couldn't start the app here")
    elif "make migrations" in query or "makemigrations" in query:
        if _run_project_command("makemigrations", cwd=CURRENT_DIR):
            _speak_and_log("Running makemigrations")
        else:
            _speak_and_log("I couldn't run makemigrations here")
    elif "migrate" in query and "makemigrations" not in query:
        if _run_project_command("migrate", cwd=CURRENT_DIR):
            _speak_and_log("Running migrate")
        else:
            _speak_and_log("I couldn't run migrate here")
    elif "what is my current folder" in query or "current folder" in query or query.strip() == "pwd":
        _speak_and_log(f"You are in {CURRENT_DIR}")
    elif query.startswith("list files") or query.startswith("show files") or query.strip() == "ls":
        try:
            items = os.listdir(CURRENT_DIR)
            if not items:
                _speak_and_log("This folder is empty.")
            else:
                preview = ", ".join(items[:20])
                _speak_and_log(f"Files here: {preview}")
        except Exception:
            _speak_and_log("I couldn't list the files here.")
    elif "change your name" in query:
        set_name()
    elif "list voices" in query or "what voices" in query:
        list_voices()
    elif "change voice" in query or "set voice" in query:
        set_voice_from_query(query)
    elif "use voice one" in query or "voice one" in query:
        set_voice_by_index(0)
        _speak_and_log("Voice set to 1")
    elif "use voice two" in query or "voice two" in query:
        set_voice_by_index(1)
        _speak_and_log("Voice set to 2")
    elif "use voice three" in query or "voice three" in query:
        set_voice_by_index(2)
        _speak_and_log("Voice set to 3")
    elif "use system voice" in query or "force system voice" in query:
        enable_system_voice()
        _speak_and_log("System voice enabled")
    elif "use engine voice" in query or "disable system voice" in query:
        enable_engine_voice()
        _speak_and_log("Engine voice enabled")
    elif "screenshot" in query:
        screenshot()
        speak("I've taken screenshot, please check it")
    elif "tell me a joke" in query:
        joke = pyjokes.get_joke()
        speak(joke)
        print(joke)
        try:
            if UI.instance is not None:
                UI.instance.log_assistant(joke)
        except Exception:
            pass
    elif "can you hear me" in query or "are you there" in query:
        _speak_and_log("Yes, I can hear you clearly.")
    elif "which model are you" in query or "what model are you" in query:
        _speak_and_log(f"I am using model {_current_model_label()}.")
    elif "audio test" in query or "test audio" in query:
        _speak_and_log("This is a voice test. If you can hear me, audio is working.")
    elif any(t in query for t in ("how are you", "bad day", "feeling down", "i feel bad", "i'm tired", "i am tired", "sad", "upset", "lonely", "stressed", "anxious", "depressed", "not okay", "burned out", "heartbroken", "i'm angry", "i am angry")):
        msg = _empathetic_response(query)
        if not msg:
            msg = "I'm here with you. Thanks for sharing that with me—how can I support you right now?"
        _speak_and_log(msg)
    elif "stop speaking" in query or "stop voice" in query or "be quiet" in query or "mute" in query:
        stop_speaking()
        _speak_and_log("Stopped speaking.")
    elif "enable startup" in query or "start on startup" in query:
        if register_startup():
            speak("I will start automatically when you log in.")
        else:
            speak("I could not enable startup on this system.")
    elif "disable startup" in query or "stop startup" in query:
        if unregister_startup():
            speak("I will not start automatically anymore.")
        else:
            speak("Startup entry was not found.")
    elif "shutdown" in query:
        speak("Shutting down the system, goodbye!")
        os.system("shutdown /s /f /t 1")
        return False
    elif "restart" in query:
        speak("Restarting the system, please wait!")
        os.system("shutdown /r /f /t 1")
        return False
    elif "offline" in query or "exit" in query or "quit" in query:
        speak("Going offline. Have a good day!")
        return False
    else:
        # Fallback to LLM for general queries
        reply = llm_generate_response(query)
        if not reply:
            reply = "Sorry, I don't have an answer for that yet."
        speak(reply)
        try:
            if UI.instance is not None:
                UI.instance.log_assistant(reply)
        except Exception:
            pass
    return True


def _listen_once_thread():
    try:
        query = takecommand()
        if query is None:
            # Give user feedback when nothing was captured
            speak("Sorry, I didn't catch that.")
            return
        cont = _handle_query(query)
        if not cont:
            try:
                if UI.instance is not None and tk is not None:
                    UI.instance.root.after(0, UI.instance.root.quit)
            except Exception:
                pass
    finally:
        try:
            if UI.instance is not None:
                UI.instance.set_status("Ready. Press N to talk.")
                UI.instance._listening_lock.release()
        except Exception:
            pass


def _wake_listener(stop_event: threading.Event) -> None:
    recognizer = sr.Recognizer()
    recognizer.dynamic_energy_threshold = True
    while not stop_event.is_set():
        if not WAKE_ENABLED:
            time.sleep(0.2)
            continue
        try:
            with sr.Microphone() as source:
                try:
                    recognizer.adjust_for_ambient_noise(source, duration=0.3)
                except Exception:
                    pass
                audio = recognizer.listen(source, timeout=2, phrase_time_limit=3)
            try:
                text = recognizer.recognize_google(audio, language="en-US").lower()
            except Exception:
                continue
            if "jarvis" in text:
                _speak_and_log("Yes, I'm listening.")
                # Immediately run a one-shot listen for the command
                _listen_once_thread()
        except Exception:
            continue


def _get_startup_folder() -> str:
    """Return the current user's Startup folder path on Windows."""
    return os.path.join(os.environ.get('APPDATA', ''),
                        'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Startup')


def register_startup(shortcut_name: str = 'Jarvis Assistant.lnk') -> bool:
    """Create a shortcut in the Startup folder so the assistant launches at login."""
    try:
        if win32com is None:
            return False
        startup = _get_startup_folder()
        if not startup or not os.path.isdir(startup):
            return False
        shell = win32com.client.Dispatch('WScript.Shell')
        shortcut_path = os.path.join(startup, shortcut_name)
        target = sys.executable
        script = os.path.abspath(__file__)
        args = f'"{script}"'
        shortcut = shell.CreateShortcut(shortcut_path)
        shortcut.TargetPath = target
        shortcut.Arguments = args
        shortcut.WorkingDirectory = os.path.dirname(script)
        shortcut.IconLocation = target
        shortcut.save()
        return True
    except Exception:
        return False


def unregister_startup(shortcut_name: str = 'Jarvis Assistant.lnk') -> bool:
    try:
        startup = _get_startup_folder()
        shortcut_path = os.path.join(startup, shortcut_name)
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)
            return True
        return False
    except Exception:
        return False


def _stdin_listener() -> None:
    """Read typed commands from stdin and pass them to the handler.

    Useful when the user types commands into the terminal like
    'open file named foo' instead of speaking them.
    """
    while True:
        try:
            raw = input().strip()
        except EOFError:
            break
        except Exception:
            continue
        if not raw:
            continue
        q = raw.lower()
        if q in ("exit", "quit"):
            break
        cont = _handle_query(q)
        if not cont:
            break


if __name__ == "__main__":
    # Ensure auto-start on Windows login (best-effort)
    try:
        registered = register_startup()
        if registered:
            print("Startup registration ensured.")
    except Exception:
        pass

    wishme()

    ui = UI()
    if ui and tk is not None:
        # Initial instructions in the transcript
        try:
            ui.log("System", "Press N or click 'Listen (N)' to speak.")
        except Exception:
            pass
        # Start wake-word listener
        wake_stop = threading.Event()
        threading.Thread(target=_wake_listener, args=(wake_stop,), daemon=True).start()
        # Also accept typed commands from terminal/stdin
        threading.Thread(target=_stdin_listener, daemon=True).start()
        ui.run()
    else:
        # Fallback to console push-to-talk: single-shot listens on Enter
        while True:
            _inp = input("Press Enter to talk (or type exit): ").strip().lower()
            if _inp in ("exit", "quit"):
                break
            q = takecommand()
            if not _handle_query(q or ""):
                break
