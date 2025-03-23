import tkinter as tk
from tkinter import ttk
import pyttsx3
import speech_recognition as sr
import sounddevice as sd
import threading
import time
import comtypes.client

# Initialize speech engine
engine = pyttsx3.init()
voices = engine.getProperty('voices')
voice_names = [voice.name for voice in voices]

# Get input devices
input_devices = [d['name'] for d in sd.query_devices() if d['max_input_channels'] > 0]

# Global flags
is_paused = False
is_stopped = False

# Use SAPI for pitch control
sapi_voice = comtypes.client.CreateObject("SAPI.SpVoice")

def speak_text(text):
    global is_paused, is_stopped
    voice_id = voices[voice_combo.current()].id
    engine.setProperty('voice', voice_id)
    rate = int(rate_entry.get())
    pitch = int(pitch_entry.get())
    engine.setProperty('rate', rate)

    while is_paused:
        time.sleep(0.1)

    if not is_stopped:
        xml_text = f'<pitch absmiddle="{pitch - 100}">{text}</pitch>'
        sapi_voice.Voice = sapi_voice.GetVoices().Item(voice_combo.current())
        sapi_voice.Rate = rate // 10 - 10
        sapi_voice.Speak(xml_text, 1)

def recognize_and_speak():
    global is_paused, is_stopped

    input_device_index = input_combo.current()
    mic_name = input_devices[input_device_index]

    recognizer = sr.Recognizer()
    try:
        device_info = sd.query_devices(mic_name)
        mic_index = device_info['index']
    except:
        log_text.insert(tk.END, f"❌ Could not find microphone: {mic_name}\n")
        return

    with sr.Microphone(device_index=mic_index) as source:
        recognizer.adjust_for_ambient_noise(source)
        log_text.insert(tk.END, "🔄 Speech recognition started. Speak now...\n")
        log_text.see(tk.END)

        is_paused = False
        is_stopped = False
        last_text = ""

        while not is_stopped:
            log_text.insert(tk.END, "🟢 Listening...\n")
            log_text.see(tk.END)

            try:
                audio = recognizer.listen(source, timeout=10, phrase_time_limit=None)
                text = recognizer.recognize_google(audio, language="ru-RU")
                text = text.strip()

                if text == last_text:
                    log_text.insert(tk.END, "⚠️ Repeated text detected. Skipping...\n")
                    log_text.see(tk.END)
                    continue

                last_text = text
                log_text.insert(tk.END, f"📝 Recognized: {text}\n")
                log_text.see(tk.END)

                speak_text(text)

            except sr.UnknownValueError:
                log_text.insert(tk.END, "🤔 Could not understand speech.\n")
                log_text.see(tk.END)
            except sr.WaitTimeoutError:
                log_text.insert(tk.END, "⌛ Speech timeout.\n")
                log_text.see(tk.END)
            except Exception as e:
                log_text.insert(tk.END, f"⚠️ Error: {e}\n")
                log_text.see(tk.END)

def start_thread():
    threading.Thread(target=recognize_and_speak).start()

def pause_speech():
    global is_paused
    is_paused = not is_paused
    status = "⏸ Paused" if is_paused else "▶️ Resumed"
    log_text.insert(tk.END, f"{status}\n")
    log_text.see(tk.END)

def stop_speech():
    global is_stopped
    is_stopped = True
    sapi_voice.Speak("", 3)
    log_text.insert(tk.END, "🛑 Speech stopped\n")
    log_text.see(tk.END)

def speak_manual_input():
    text = manual_entry.get("1.0", tk.END)
    if text.strip():
        speak_text(text.strip())
        log_text.insert(tk.END, f"📢 Spoken manually: {text}\n")
        log_text.see(tk.END)

# GUI
root = tk.Tk()
try:
    root.iconbitmap("icon.ico")
except:
    pass  # Иконка не обязательна, не рушим запуск

root.title("SpeechMachine")
root.configure(bg="#16151B")

style = ttk.Style()
style.theme_use("default")
style.configure("TLabel", background="#16151B", foreground="white")
style.configure("TButton", background="#3E3A57", foreground="white", padding=6)
style.configure("TCombobox", fieldbackground="black", background="black", foreground="white")

def combo_config(combo):
    combo.configure(background="black", foreground="white")

# Interface Elements
ttk.Label(root, text="🗣 Voice:").grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)
voice_combo = ttk.Combobox(root, values=voice_names, width=40)
voice_combo.grid(column=1, row=0, padx=5, pady=5)
voice_combo.current(0)

ttk.Label(root, text="🎤 Microphone:").grid(column=0, row=1, sticky=tk.W, padx=5, pady=5)
input_combo = ttk.Combobox(root, values=input_devices, width=40)
input_combo.grid(column=1, row=1, padx=5, pady=5)
input_combo.current(0)

ttk.Label(root, text="🚀 Speech Rate (50–300):").grid(column=0, row=3, sticky=tk.W, padx=5, pady=5)
rate_entry = ttk.Entry(root, width=10)
rate_entry.insert(0, "100")
rate_entry.grid(column=1, row=3, padx=5, pady=5, sticky=tk.W)

ttk.Label(root, text="🎵 Pitch (0–200):").grid(column=0, row=4, sticky=tk.W, padx=5, pady=5)
pitch_entry = ttk.Entry(root, width=10)
pitch_entry.insert(0, "100")
pitch_entry.grid(column=1, row=4, padx=5, pady=5, sticky=tk.W)

ttk.Label(root, text="✍️ Enter text manually:").grid(column=0, row=5, sticky=tk.W, padx=5, pady=5)
manual_entry = tk.Text(root, height=5, width=60, bg="black", fg="white", insertbackground="white")
manual_entry.grid(column=1, row=5, padx=5, pady=5, sticky="we")
manual_speak_button = ttk.Button(root, text="🔊 Speak", command=speak_manual_input)
manual_speak_button.grid(column=1, row=6, padx=5, pady=5, sticky=tk.W)

log_text = tk.Text(root, height=10, width=60, bg="black", fg="white", insertbackground="white")
log_text.grid(column=0, row=7, columnspan=2, padx=5, pady=10)

button_frame = ttk.Frame(root)
button_frame.grid(column=0, row=8, columnspan=2, pady=10)

ttk.Button(button_frame, text="🎙️ Start", command=start_thread).grid(column=0, row=0, padx=5)
ttk.Button(button_frame, text="⏸ Pause", command=pause_speech).grid(column=1, row=0, padx=5)
ttk.Button(button_frame, text="🛑 Stop", command=stop_speech).grid(column=2, row=0, padx=5)

root.mainloop()