import speech_recognition as sr
import webbrowser
import datetime
import pyttsx3
import time
import os
import subprocess
import random
import sys
import pyautogui
from difflib import get_close_matches

# Try to import optional packages
try:
    import screen_brightness_control as sbc
    BRIGHTNESS_CONTROL = True
except ImportError:
    BRIGHTNESS_CONTROL = False
    print("Note: Brightness control disabled (missing screen_brightness_control package)")

class VoiceAssistant:
    def __init__(self):
        # Initialize components
        self.engine = pyttsx3.init()
        self.recognizer = sr.Recognizer()
        self.recognizer.energy_threshold = 400
        self.command_history = []
        
        # Configure voice with more personality
        voices = self.engine.getProperty('voices')
        self.engine.setProperty('voice', voices[1].id)  # Female voice
        self.engine.setProperty('rate', 165)  # Slightly slower for clarity
        self.engine.setProperty('volume', 1.0)  # Maximum volume

        # Enhanced website list
        self.website_commands = {
            'youtube': 'https://youtube.com',
            'google': 'https://google.com',
            'github': 'https://github.com',
            'wikipedia': 'https://wikipedia.org',
            'amazon': 'https://amazon.com',
            'netflix': 'https://netflix.com',
            'spotify': 'https://spotify.com',
            'twitter': 'https://twitter.com',
            'facebook': 'https://facebook.com',
            'instagram': 'https://instagram.com',
            'reddit': 'https://reddit.com',
            'linkedin': 'https://linkedin.com',
            'outlook': 'https://outlook.com',
            'gmail': 'https://mail.google.com',
            'drive': 'https://drive.google.com',
            'maps': 'https://maps.google.com',
            'weather': 'https://weather.com',
            'news': 'https://news.google.com',
            'imdb': 'https://imdb.com',
            'stack overflow': 'https://stackoverflow.com'
        }

    def speak(self, text, is_priority=False, response_type=None):
        """Enhanced text-to-speech with more personality"""
        responses = {
            'acknowledge': [
                "Right away, sir.", 
                "On it, sir.",
                "Processing your request now.",
                "I'll take care of that for you.",
                "Working on that immediately."
            ],
            'success': [
                "All done, sir.",
                "Task completed successfully.",
                "I've taken care of that for you.",
                "Mission accomplished.",
                "Your request has been fulfilled."
            ],
            'error': [
                "I'm sorry, I couldn't complete that request.",
                "Apologies, sir. I encountered an issue.",
                "I wasn't able to do that for you.",
                "Something went wrong with that command.",
                "My systems report a problem with that request."
            ],
            'greeting': [
                "At your service, sir. How may I assist you today?",
                "Ready and awaiting your commands.",
                "How can I be of assistance today?",
                "I'm here to help. What would you like me to do?"
            ]
        }
        
        if response_type in responses:
            text = f"{random.choice(responses[response_type])} {text}"
        
        print(f"Friday: {text}")
        if is_priority:
            self.engine.stop()
        
        # Split into sentences for more natural speech
        sentences = [s.strip() for s in text.split('.') if s.strip()]
        for sentence in sentences:
            self.engine.say(sentence)
            self.engine.runAndWait()
            time.sleep(0.15)  # Natural pause between sentences

    def greet(self):
        """More personalized time-based greeting"""
        hour = datetime.datetime.now().hour
        day_part = "morning" if 5 <= hour < 12 else "afternoon" if 12 <= hour < 17 else "evening"
        
        greetings = {
            'morning': [
                "A very good morning to you, sir! How may I assist you today?",
                "Morning, sir! Ready to tackle the day together?",
                "Good morning! What would you like me to help with today?"
            ],
            'afternoon': [
                "Good afternoon, sir. How can I be of service?",
                "Afternoon protocols engaged. What's on the agenda?",
                "Good afternoon! What would you like me to do for you?"
            ],
            'evening': [
                "Good evening, sir. How may I assist you tonight?",
                "Evening systems online. Your commands, sir?",
                "Good evening! What can I do for you at this hour?"
            ]
        }
        
        self.speak(random.choice(greetings[day_part]), response_type='greeting')

    def recognize_speech(self, timeout=5):
        """Enhanced speech recognition with more feedback"""
        with sr.Microphone() as source:
            try:
                print("\n[System] Listening... (speak now)")
                self.speak("I'm listening, sir.", response_type='acknowledge')
                self.recognizer.adjust_for_ambient_noise(source, duration=0.8)
                audio = self.recognizer.listen(source, timeout=timeout, phrase_time_limit=8)
                
                command = self.recognizer.recognize_google(audio).lower()
                print(f"You said: {command}")
                self.speak("Command received and understood.", response_type='acknowledge')
                return command
                
            except sr.WaitTimeoutError:
                self.speak("I didn't hear anything. Shall I continue listening?", response_type='error')
                return None
            except sr.UnknownValueError:
                self.speak("I couldn't understand that. Could you please repeat?", response_type='error')
                return None
            except Exception as e:
                self.speak(f"Audio processing error: {str(e)}", response_type='error')
                return None

    def execute_command(self, command):
        """Enhanced command execution with more conversational responses"""
        if not command:
            self.speak("I didn't receive any command, sir.", response_type='error')
            return
            
        self.command_history.append({
            'time': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'command': command
        })
        
        # Application commands
        app_commands = {
            'notepad': ('notepad.exe', "Opening your notes", "I couldn't open Notepad"),
            'calculator': ('calc.exe', "Launching the calculator", "Calculator access failed"),
            'chrome': ('chrome.exe', "Opening Chrome browser", "Couldn't launch Chrome"),
            'spotify': ('spotify.exe', "Starting Spotify for you", "Spotify couldn't be opened"),
            'word': ('WINWORD.EXE', "Opening Microsoft Word", "Word document failed to open"),
            'excel': ('EXCEL.EXE', "Launching Excel spreadsheet", "Excel couldn't be started")
        }
        
        for app, (exe, success_msg, error_msg) in app_commands.items():
            if f'open {app}' in command:
                try:
                    self.speak(success_msg, response_type='acknowledge')
                    os.startfile(exe)
                    self.speak(f"{app.capitalize()} is now ready for you.", response_type='success')
                    return
                except Exception:
                    self.speak(error_msg, response_type='error')
                    return
        
        # Website commands with more feedback
        for site, url in self.website_commands.items():
            if f'open {site}' in command:
                self.speak(f"Accessing {site.replace('_', ' ')} for you.", response_type='acknowledge')
                webbrowser.open(url)
                self.speak(f"{site.capitalize()} is now available in your browser.", response_type='success')
                return
        
        # System commands with confirmation
        if 'shutdown' in command:
            self.speak("Initiating system shutdown sequence. Confirm?", response_type='acknowledge')
            confirm = self.recognize_speech()
            if confirm and 'yes' in confirm:
                self.speak("Shutting down now. Goodbye, sir.", response_type='acknowledge')
                os.system("shutdown /s /t 1")
            else:
                self.speak("Shutdown cancelled.", response_type='success')
        elif 'restart' in command:
            self.speak("Preparing to restart the system. Should I proceed?", response_type='acknowledge')
            confirm = self.recognize_speech()
            if confirm and 'yes' in confirm:
                self.speak("Rebooting now. See you soon, sir.", response_type='acknowledge')
                os.system("shutdown /r /t 1")
            else:
                self.speak("Restart aborted.", response_type='success')
        elif 'lock' in command:
            self.speak("Securing your workstation now.", response_type='acknowledge')
            os.system("rundll32.exe user32.dll,LockWorkStation")
        
        # Enhanced special commands
        elif 'time' in command:
            current_time = datetime.datetime.now()
            hour = current_time.hour
            minute = current_time.minute
            am_pm = "AM" if hour < 12 else "PM"
            hour_12 = hour if hour <= 12 else hour - 12
            
            time_greetings = [
                f"The current time is {hour_12}:{minute:02d} {am_pm}",
                f"My systems show the time as {hour_12}:{minute:02d} {am_pm}",
                f"It's now {hour_12}:{minute:02d} {am_pm}",
                f"The clock reads {hour_12}:{minute:02d} {am_pm}"
            ]
            
            if hour < 5:
                time_greetings.append(f"It's {hour_12}:{minute:02d} {am_pm}. Quite late, sir. Shouldn't you be resting?")
            elif hour < 12:
                time_greetings.append(f"Good morning! The time is {hour_12}:{minute:02d} {am_pm}")
            elif hour < 17:
                time_greetings.append(f"Good afternoon! It's {hour_12}:{minute:02d} {am_pm}")
            else:
                time_greetings.append(f"Good evening! The time is now {hour_12}:{minute:02d} {am_pm}")
            
            self.speak(random.choice(time_greetings), response_type='success')

        elif 'date' in command:
            current_date = datetime.datetime.now()
            weekday = current_date.strftime("%A")
            month = current_date.strftime("%B")
            day = current_date.strftime("%d").lstrip('0')
            year = current_date.strftime("%Y")
            
            date_responses = [
                f"Today is {weekday}, {month} {day}",
                f"The date today is {weekday}, {month} {day}",
                f"According to my calendar, it's {weekday}, {month} {day}",
                f"We're currently on {weekday}, {month} {day}"
            ]
            
            if 'year' in command:
                date_responses.append(f"The year is {year}")
                date_responses.append(f"We're in the year {year}")
            
            for response in date_responses:
                self.speak(response, response_type='success')
                time.sleep(0.3)

        elif 'search' in command:
            query = command.replace('search', '').replace('for', '').strip()
            if query:
                search_phrases = [
                    f"Searching the web for information about {query}",
                    f"Looking up {query} for you",
                    f"Let me find results about {query}",
                    f"Accessing search engines for {query}"
                ]
                self.speak(random.choice(search_phrases), response_type='acknowledge')
                webbrowser.open(f"https://google.com/search?q={query}")
                self.speak(f"I've displayed search results for {query} in your browser.", response_type='success')
            else:
                self.speak("What would you like me to search for? Please be specific.", response_type='error')

        elif 'brightness' in command and BRIGHTNESS_CONTROL:
            current = sbc.get_brightness()[0]
            if 'up' in command:
                new = min(100, current + 20)
                sbc.set_brightness(new)
                self.speak(f"Screen brightness increased to {new} percent. Is this comfortable?", response_type='success')
            elif 'down' in command:
                new = max(0, current - 20)
                sbc.set_brightness(new)
                self.speak(f"Brightness decreased to {new} percent. Better for your eyes now?", response_type='success')
            elif 'set' in command:
                try:
                    level = int(command.split()[-1])
                    if 0 <= level <= 100:
                        sbc.set_brightness(level)
                        self.speak(f"Screen brightness adjusted to {level} percent as requested.", response_type='success')
                    else:
                        self.speak("Please specify a brightness level between 0 and 100 percent.", response_type='error')
                except ValueError:
                    self.speak("I didn't understand the brightness level you requested.", response_type='error')

        elif 'exit' in command or 'quit' in command or 'shutdown' in command:
            farewells = [
                "Shutting down systems now. Goodbye, sir.",
                "Powering off. It was a pleasure assisting you.",
                "Disconnecting. Have a wonderful day, sir!",
                "Friday signing off. Until next time!"
            ]
            self.speak(random.choice(farewells), response_type='acknowledge')
            sys.exit()

        else:
            suggestions = [
                "I'm not sure I understood that command. Try saying things like:",
                "My apologies, I didn't catch that. You can ask me to:",
                "I didn't recognize that request. Here are some things I can do:"
            ]
            
            examples = [
                "Open applications like Notepad or Chrome",
                "Visit websites like YouTube or Wikipedia",
                "Tell you the current time or date",
                "Search the web for information",
                "Adjust your screen brightness",
                "Lock your computer or restart it"
            ]
            
            self.speak(random.choice(suggestions), response_type='error')
            time.sleep(0.5)
            for example in random.sample(examples, 3):
                self.speak(example)
                time.sleep(0.3)

    def run(self):
        """Main execution loop with enhanced interaction"""
        self.greet()
        while True:
            try:
                # Check for inactivity
                if len(self.command_history) > 0:
                    last_cmd_time = datetime.datetime.strptime(self.command_history[-1]['time'], "%Y-%m-%d %H:%M:%S")
                    if (datetime.datetime.now() - last_cmd_time).seconds > 120:
                        self.speak("I'm still here if you need me, sir.", response_type='greeting')
                
                command = self.recognize_speech()
                if command:
                    self.execute_command(command)
                
                time.sleep(0.5)
                
            except KeyboardInterrupt:
                self.speak("Emergency shutdown initiated. Goodbye, sir!", response_type='error')
                sys.exit()
            except Exception as e:
                self.speak(f"System error detected: {str(e)}. Attempting to recover...", response_type='error')
                time.sleep(1)
                continue

if __name__ == "__main__":
    assistant = VoiceAssistant()
    assistant.run()