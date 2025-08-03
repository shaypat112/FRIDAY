import speech_recognition as sr
import webbrowser
import datetime
import pyttsx3
import time
import os
import subprocess
import random
import sys
import ctypes
import screen_brightness_control as sbc
import pyautogui

# Initialize components
engine = pyttsx3.init()
recognizer = sr.Recognizer()

# Configure voice (Friday-style)
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)  # Try 0 or 2 for different voices
engine.setProperty('rate', 170)  # Slightly faster pace
engine.setProperty('volume', 0.8)  # Slightly louder

# Application paths (customize these to match your system)
APP_PATHS = {
    'whatsapp': r"C:\Users\shiva\AppData\Local\WhatsApp\WhatsApp.exe",
    'spotify': r"C:\Users\shiva\AppData\Roaming\Spotify\Spotify.exe",
    'vscode': r"C:\Users\shiva\AppData\Local\Programs\Microsoft VS Code\Code.exe",
    'chrome': r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    'notepad': r"C:\Windows\System32\notepad.exe",
    'calculator': r"C:\Windows\System32\calc.exe",
    'word': r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    'excel': r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
    'powerpoint': r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
    'outlook': r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
}

# Website URLs
WEBSITES = {
    'youtube': 'https://youtube.com',
    'gmail': 'https://mail.google.com',
    'google': 'https://google.com',
    'github': 'https://github.com',
    'linkedin': 'https://linkedin.com',
    'twitter': 'https://twitter.com',
    'facebook': 'https://facebook.com',
    'instagram': 'https://instagram.com',
    'amazon': 'https://amazon.com',
    'netflix': 'https://netflix.com',
    'wikipedia': 'https://wikipedia.org',
    'reddit': 'https://reddit.com',
    'stackoverflow': 'https://stackoverflow.com',
    'discord': 'https://discord.com',
    'zoom': 'https://zoom.us'
}

# System commands
SYSTEM_COMMANDS = {
    'shutdown': 'shutdown /s /t 1',
    'restart': 'shutdown /r /t 1',
    'sleep': 'rundll32.exe powrprof.dll,SetSuspendState 0,1,0',
    'lock': 'rundll32.exe user32.dll,LockWorkStation'
}

# Conversation phrases
RESPONSES = {
    'acknowledge': [
        "Right away, sir.",
        "On it, sir.",
        "Processing that now.",
        "Executing your command."
    ],
    'success': [
        "Task completed successfully.",
        "Done, sir.",
        "Operation finished.",
        "Your request has been fulfilled."
    ],
    'error': [
        "I encountered a problem with that, sir.",
        "Apologies, I couldn't complete that.",
        "There seems to be an issue.",
        "My systems report a malfunction."
    ]
}

def speak(text, priority=False, type=None):
    """Enhanced Friday-style text-to-speech with conversational responses"""
    if type and type in RESPONSES:
        text = random.choice(RESPONSES[type]) + " " + text
    
    print(f"Friday: {text}")
    if priority:
        engine.stop()  # Stop any current speech
    engine.say(text)
    engine.runAndWait()

def greet():
    """Personalized greeting with conversational tone"""
    hour = datetime.datetime.now().hour
    
    greeting_options = {
        'morning': [
            f"Good morning, sir. How may I assist you today?",
            f"Morning, sir. Ready for your commands."
        ],
        'afternoon': [
            f"Good afternoon Sir, Protocols Initiated. What can I do for you?",
            f"Afternoon, How may I be of service?"
        ],
        'evening': [
            f"Good evening, Protocols Initiated. Your commands?",
            f"Evening, sir. Awaiting your instructions."
        ]
    }
    
    if 5 <= hour < 12:
        greeting = random.choice(greeting_options['morning'])
    elif 12 <= hour < 18:
        greeting = random.choice(greeting_options['afternoon'])
    else:
        greeting = random.choice(greeting_options['evening'])
    
    speak(greeting)

def recognize_speech(timeout=5):
    """Enhanced speech recognition with conversational feedback"""
    with sr.Microphone() as source:
        for attempt in range(3):
            try:
                speak("I'm listening, sir.", type='acknowledge')
                print("\n[System] Listening...")
                recognizer.adjust_for_ambient_noise(source)
                audio = recognizer.listen(source, timeout=timeout)
                
                command = recognizer.recognize_google(audio).lower()
                print(f"You said: {command}")
                speak("Command received.", type='acknowledge')
                return command
                
            except sr.WaitTimeoutError:
                if attempt < 2:
                    speak("I didn't catch that, sir. Please repeat.", type='error')
                    continue
                return None
            except Exception as e:
                speak(f"Audio processing error: {str(e)}", type='error')
                return None
        return None

def execute_command(command):
    """Enhanced command execution with full conversational feedback"""
    if not command:
        speak("No command detected, sir.", type='error')
        return
        
    # Application commands
    for app_name, app_path in APP_PATHS.items():
        if f'open {app_name}' in command:
            if os.path.exists(app_path):
                speak(f"Initializing {app_name} application.", type='acknowledge')
                subprocess.Popen([app_path])
                speak(f"{app_name.capitalize()} is now running, sir.", type='success')
            else:
                speak(f"I couldn't locate the {app_name} application.", type='error')
            return
    
    # Website commands
    for site_name, site_url in WEBSITES.items():
        if f'open {site_name}' in command:
            speak(f"Accessing {site_name} servers.", type='acknowledge')
            webbrowser.open(site_url)
            speak(f"{site_name.capitalize()} is now available on your display, sir.", type='success')
            return
    
    # System commands
    for sys_cmd, cmd_string in SYSTEM_COMMANDS.items():
        if sys_cmd in command:
            speak(f"Initiating system {sys_cmd} sequence.", type='acknowledge')
            os.system(cmd_string)
            return
    
    # Special commands
    if 'time' in command:
        current_time = datetime.datetime.now().strftime("%I:%M %p")
        speak(f"Current system time is {current_time}, sir.", type='success')
    
    elif 'date' in command:
        current_date = datetime.datetime.now().strftime("%A, %B %d, %Y")
        speak(f"Today's date is {current_date}, sir.", type='success')
    
    elif 'search' in command:
        query = command.replace('search', '').strip()
        if query:
            speak(f"Initiating search protocol for: {query}", type='acknowledge')
            webbrowser.open(f"https://google.com/search?q={query}")
            speak(f"Search results for {query} are now displayed, sir.", type='success')
        else:
            speak("Please specify your search query, sir.", type='error')
    
    elif 'brightness up' in command:
        current = sbc.get_brightness()[0]
        new = min(100, current + 20)
        sbc.set_brightness(new)
        speak(f"Screen brightness increased to {new} percent, sir.", type='success')
    
    elif 'brightness down' in command:
        current = sbc.get_brightness()[0]
        new = max(0, current - 20)
        sbc.set_brightness(new)
        speak(f"Screen brightness decreased to {new} percent, sir.", type='success')
    
    elif 'volume up' in command:
        pyautogui.press('volumeup')
        speak("Volume increased, sir.", type='success')
    
    elif 'volume down' in command:
        pyautogui.press('volumedown')
        speak("Volume decreased, sir.", type='success')
    
    elif 'mute' in command:
        pyautogui.press('volumemute')
        speak("Audio muted, sir.", type='success')
    
    elif 'take screenshot' in command:
        screenshot = pyautogui.screenshot()
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"screenshot_{timestamp}.png"
        screenshot.save(filename)
        speak(f"Screenshot saved as {filename}, sir.", type='success')
    
    elif 'exit' in command or 'quit' in command or 'shutdown' in command:
        speak("Initiating shutdown sequence. Goodbye, sir.", type='acknowledge')
        sys.exit()
    
    else:
        speak("Command syntax not recognized. Please rephrase, sir.", type='error')

def display_help():
    """Display available commands"""
    print("\nAvailable Commands:")
    print("\nApplications:")
    for app in APP_PATHS.keys():
        print(f"- 'open {app}'")
    
    print("\nWebsites:")
    for site in WEBSITES.keys():
        print(f"- 'open {site}'")
    
    print("\nSystem Commands:")
    for cmd in SYSTEM_COMMANDS.keys():
        print(f"- '{cmd}'")
    
    print("\nOther Commands:")
    print("- 'time' (current time)")
    print("- 'date' (current date)")
    print("- 'search [query]' (web search)")
    print("- 'brightness up/down'")
    print("- 'volume up/down/mute'")
    print("- 'take screenshot'")
    print("")
    

def main():
    """Main assistant loop with enhanced conversation"""
    greet()
    display_help()
    
    while True:
        # Conversational prompt
        speak("Awaiting your next command, sir.")
        
        # Get and execute command
        command = recognize_speech()
        if command:
            execute_command(command)
        
        # Small delay to prevent CPU overuse
        time.sleep(0.5)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        speak("Emergency shutdown protocol activated. Goodbye, sir.", type='error')
    except Exception as e:
        speak(f"Critical system failure detected: {str(e)}. Rebooting...", type='error')
        main()