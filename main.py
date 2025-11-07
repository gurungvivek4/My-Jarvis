from win32com.client import Dispatch          # For Windows speech output
import speech_recognition as sr               # For speech to text
import webbrowser                             # For opening websites
import os                                     # For system commands
import pyjokes                                # For jokes
import datetime                               # For time queries
from openai import OpenAI                     # For ChatGPT API
import music_library                          # Custom music library
from config import apikey                     # <-- CHANGE 1: Import the key

# ---------------------- INITIAL SETUP -------------------------
speak = Dispatch("SAPI.SpVoice")              # Windows voice engine
r = sr.Recognizer()                           # Speech recognizer
open_site = webbrowser.open                   # Shortcut for webbrowser

# ---------------------- OPENAI API KEY ------------------------
client = OpenAI(api_key=apikey)

# ---------------------- WEBSITE SHORTCUTS ---------------------
sites = {
    "youtube": "https://www.youtube.com/",
    "chat gpt": "https://chatgpt.com/",
    "google": "https://www.google.com/",
    "wikipedia": "https://www.wikipedia.org/",
    "github": "https://github.com/",
    "linkedin": "https://www.linkedin.com/"
}

# ---------------------- LISTEN FUNCTION -----------------------
def listen():
    """Listens to the user's voice and converts it into text."""
    with sr.Microphone() as source:
        print("\n🎧 Listening...")
        r.adjust_for_ambient_noise(source, duration=1)
        try:
            audio = r.listen(source, timeout=5, phrase_time_limit=6)
        except Exception as e:
            print("❌ Listen error:", e)
            return ""

    try:
        text = r.recognize_google(audio, language="en-in").lower()
        print("🗣️ You said:", text)
        speak.Speak(f"I heard you say {text}")
        return text
    except sr.UnknownValueError:
        speak.Speak("Sorry, I didn't catch that. Please repeat.")
        return ""
    except sr.RequestError:
        speak.Speak("Network issue detected. Try again later.")
        return ""
    except Exception as e:
        print("❌ Recognition error:", e)
        return ""

# ---------------------- CHAT WITH GPT -------------------------
def ask_gpt(question):
    """Ask OpenAI a question and speak the response."""
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are Jarvis, a smart and witty assistant that helps Vivek."},
                {"role": "user", "content": question}
            ]
        )
        answer = response.choices[0].message.content
        print("🤖 ChatGPT:", answer)
        speak.Speak(answer)
    except Exception as e:
        print("❌ OpenAI Error:", e)
        speak.Speak("Sorry, I'm unable to connect to OpenAI right now.")

# ---------------------- COMMAND FUNCTIONS ---------------------
def open_website(text):
    for name, url in sites.items():
        if name in text:
            speak.Speak(f"Opening {name}")
            open_site(url)
            return True
    return False

def play_music(text):
    for name, song in music_library.music.items():
        if name in text:
            speak.Speak(f"Playing {name}")
            open_site(song)
            return True
    speak.Speak("Sorry, I couldn't find that song.")
    return False

def tell_joke():
    joke = pyjokes.get_joke()
    print("😂", joke)
    speak.Speak(joke)

def tell_time():
    now = datetime.datetime.now().strftime("%I:%M %p")
    speak.Speak(f"The time is {now}")

def self_quit():
    speak.Speak("As you command, Vivek. Shutting myself down now.")
    os._exit(0)

# ---------------------- STARTUP MESSAGE -----------------------
speak.Speak("Jarvis version 4 point 3 activated and ready for your work, Vivek!")

# ---------------------- MAIN LOOP -----------------------------
while True:
    text = listen()
    if not text:
        continue

    # --- SELF-QUIT COMMANDS ---
    if any(phrase in text for phrase in [
        "jarvis quit", "jarvis quit yourself", "jarvis shut down",
        "jarvis terminate yourself", "jarvis power off"
    ]):
        self_quit()

    # --- NORMAL EXIT PHRASES (DOES NOT QUIT) ---
    elif any(word in text for word in ["stop", "exit", "bye", "goodbye"]):
        speak.Speak("I'm still here, Vivek. Say 'Jarvis quit yourself' if you want me to shut down.")

    # --- GREETING COMMANDS ---
    elif any(word in text for word in ["hello", "hi jarvis", "hey jarvis", "good morning", "good evening"]):
        speak.Speak("Hello Vivek! How are you doing today?")

    # --- OPEN WEBSITE COMMAND ---
    elif "open" in text:
        if not open_website(text):
            speak.Speak("Sorry, I don't recognize that website.")

    # --- MUSIC COMMAND ---
    elif "play" in text:
        play_music(text)

    # --- TIME COMMAND ---
    elif "time" in text:
        tell_time()

    # --- JOKE COMMAND ---
    elif "joke" in text:
        tell_joke()

    # --- ASK CHATGPT / AI QUESTION ---
    elif any(word in text for word in ["ask", "chatgpt", "question", "explain", "define"]):
        speak.Speak("What would you like to ask me?")
        question = listen()
        if question:
            ask_gpt(question)
        else:
            speak.Speak("I didn't hear your question, Vivek.")