import sys
import os
import pyautogui
import winsound
import threading
import pygetwindow as gw
import time
from groq import Groq
import speech_recognition as sr
import pyttsx3
import webbrowser, docx
import subprocess
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, \
    QHBoxLayout, QPushButton, QLineEdit, QTextEdit, QLabel, QRadioButton
from PyQt5.QtGui import QFont, QIcon
from office import open_word_document, close_word_document, save_document
from datetime import datetime
import queue
from app_opener import open_application, take_screenshot
from dotenv import load_dotenv
from PyQt5.QtGui import QPixmap, QMovie
from PyQt5.QtCore import Qt, QSize


# =================FILE OPENING FUNCTIONS=====================

def open_file_by_name(filename):
    filename = filename.lower().strip()

    words = filename.replace(".", " ").replace("_", " ").split()

    extensions = ["pdf", "docx", "txt", "xlsx", "pptx"]
    file_ext = None

    for ext in extensions:
        if ext in words:
            file_ext = ext
            words.remove(ext)

    keywords = words

    search_dirs = [
    os.path.join(os.path.expanduser("~"), "Desktop"),
    os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop"),  # ✅ ADD THIS
    os.path.join(os.path.expanduser("~"), "Downloads"),
    os.path.join(os.path.expanduser("~"), "Documents"),
    "C:\\",
    "D:\\",
    "E:\\",
]

    best_match = None
    best_score = 0

    for folder in search_dirs:
        for root, dirs, files in os.walk(folder):
            # ❌ Skip heavy/system folders
            dirs[:] = [d for d in dirs if d.lower() not in [
                        "windows", "program files", "program files (x86)", "appdata", "$recycle.bin"
                        ]]
            for file in files:
                file_lower = file.lower()

                # Extension filter
                if file_ext and not file_lower.endswith("." + file_ext):
                    continue

                # Count keyword matches
                matched_words = [word for word in keywords if word in file_lower]
                score = len(matched_words)

                # 🔹 IMPORTANT: ignore weak matches
                if len(keywords) > 1 and score < 2:
                    continue

                if score > best_score:
                    best_score = score
                    best_match = os.path.join(root, file)

    if best_match:
        os.startfile(best_match)
        return best_match

    return None




            
            
def install_app_dynamic(app_query):
    try:
        app_query = app_query.lower().replace("install", "").replace("download", "").strip()

        command = f'winget install --name "{app_query}" --accept-source-agreements --accept-package-agreements'

        result = subprocess.run(command, shell=True, capture_output=True, text=True)

        if "No package found" in result.stdout:
            # 🔹 fallback to Microsoft Store
            os.startfile(f"ms-windows-store://search/?query={app_query}")
            return f"{app_query} not found in winget, opening Microsoft Store"

        return app_query

    except Exception as e:
        print("Install error:", e)
        return None
    
    

DOWNLOADS_FOLDER = os.path.join(os.path.expanduser("~"), "Downloads")

class ChatWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.typing_mode = None
        self.notepad_process = None
        self.notepad_file_path = None
        self.tts_queue = queue.Queue()
        # ✅ CREATE ENGINE FIRST
        self.tts_engine = pyttsx3.init()
        # ✅ THEN START THREAD
        threading.Thread(target=self._tts_worker, daemon=True).start()
        self.notepad_open = False
        self.init_ui()
        self.word_app = None
        self.doc = None
        self.listen_for_text = False
        self.continued_input = False
        self.last_response = ""
        self.voice_gender = "male"
        self.manual_path = "commands_manual.pdf"
        #changes i have made to the code
        self.open_apps = {}
        self.current_app = None

    def init_ui(self):
        groq_api_key = "your_api_key_here"
        self.client = Groq(api_key=groq_api_key)

        self.setWindowTitle("Cypher Voice AI Assistant")  # Set window title
        self.setGeometry(100, 100, 600, 400)

        self.setWindowIcon(QIcon('va.ico'))

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout()
        central_widget.setLayout(layout)

        text_label = QLabel("How can I help?")
        text_label.setStyleSheet("color: #66FCF1; font-size: 24px;")
        text_label.setAlignment(Qt.AlignCenter)  # Align text to the center
        layout.addWidget(text_label)

        
        # Input box
        input_layout = QHBoxLayout()
        layout.addLayout(input_layout)

        input_label = QLabel("Input:")
        input_label.setStyleSheet("color: white;")
        input_layout.addWidget(input_label)

        self.user_input = QLineEdit()
        self.user_input.setStyleSheet("background-color: #022140; color: white;")  #input box
        input_layout.addWidget(self.user_input)

        
        self.chat_history_display = QTextEdit()
        self.chat_history_display.setStyleSheet("background-color: #022140; color: white;") 
        self.chat_history_display.setReadOnly(True)  
        layout.addWidget(self.chat_history_display)

        
        self.chat_history_display.append("Hello, I'm Cypher. How can I assist you today?")

        # GIF image
        movie = QMovie("mic.gif")  # Replace "wave.gif" with your GIF path
        movie.setScaledSize(QSize(600, 200))  # Adjust the size of the GIF
        movie.start()
        gif_label = QLabel()
        gif_label.setMovie(movie)
        layout.addWidget(gif_label)




        # Send button
        send_button = QPushButton("Send")
        send_button.setStyleSheet("background-color: #4681f4; color: white;")  
        layout.addWidget(send_button)
        send_button.clicked.connect(self.send_message)










        # Voice input button
        voice_button = QPushButton("Voice Input")
        voice_button.setStyleSheet("background-color: #5783db; color: white !important;")  
        layout.addWidget(voice_button)
        voice_button.clicked.connect(self.listen_voice_input)

        read_aloud_button = QPushButton("Read Aloud")
        read_aloud_button.setStyleSheet("background-color: #55c2da; color: white;")   
        layout.addWidget(read_aloud_button)
        read_aloud_button.clicked.connect(self.read_document_aloud)

        self.listening_label = QLabel()
        self.listening_label.setStyleSheet("color: #9AC8CD;")
        layout.addWidget(self.listening_label)

        gender_layout = QHBoxLayout()
        male_radio = QRadioButton("Male")
        male_radio.setChecked(True)  # Default to male
        male_radio.setStyleSheet("color: #A3D8FF;")
        male_radio.toggled.connect(lambda: self.set_voice_gender("male"))
        female_radio = QRadioButton("Female")
        female_radio.setStyleSheet("color: #A3D8FF;")
        female_radio.toggled.connect(lambda: self.set_voice_gender("female"))
        gender_layout.addWidget(male_radio)
        gender_layout.addWidget(female_radio)
        layout.addLayout(gender_layout)
        self.setStyleSheet("background-color: #1F2833;")  # Set main window background color to white
        self.show()




    def _tts_worker(self):
        while True:
            text = self.tts_queue.get()
            if text is None:
                break

            try:
                if self.voice_gender == "male":
                    self.tts_engine.setProperty(
                    'voice',
                    'HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Speech\\Voices\\Tokens\\TTS_MS_EN-US_DAVID_11.0'
                )
                else:
                    self.tts_engine.setProperty(
                    'voice',
                    'HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Speech\\Voices\\Tokens\\TTS_MS_EN-US_ZIRA_11.0'
                )

                clean_text = text.replace("**", "").replace("\n", ". ")

                # 🔹 ADD THIS LINE HERE
                clean_text = clean_text[:1500]
                chunks = [clean_text[i:i+300] for i in range(0, len(clean_text), 300)]
                # 🔹 IMPORTANT: reset engine before speaking
                self.tts_engine.stop()
                
                for chunk in chunks:
                    self.tts_engine.say(chunk)
                self.tts_engine.runAndWait()
                time.sleep(0.2)

            except Exception as e:
                print("TTS error:", e)
            
    def send_message(self):
        query = self.user_input.text().strip().lower()  # Convert the query to lowercase and strip whitespaces
        self.user_input.clear()
        self.update_chat_history(f"User: {query}\n")

        if query == "quit":
            self.close()
            return
        
        if query == "user manual" or query == "commands manual" :
            print("Opening user manual")
            self.display_help()
            return
        
        
        # =====================================OPEN NOTEPAD=========================================
        if query == "open notepad":
            subprocess.Popen(["notepad.exe"])
            time.sleep(2)
            self.notepad_open = True
            self.typing_mode = "notepad"
            self.update_chat_history("Notepad opened.\n")
            self.read_out_loud("Notepad opened.")
            return
        
        
        if query == "start typing in notepad":
            if self.notepad_open:
                self.listen_for_text = True
                self.typing_mode = "notepad"
                self.update_chat_history("Listening for typing in notepad....")
                self.showMinimized()   # 👈 add this
                self.read_out_loud("Please start speaking.")
            else:
                self.update_chat_history("Please open notepad first.\n")
            return
        
        if query == "close notepad":
            subprocess.Popen(["taskkill", "/F", "/IM", "notepad.exe"], shell=True)
            self.notepad_open = False
            self.update_chat_history("Notepad closed.\n")
            self.read_out_loud("Notepad closed.")
            return
        
        if query == "stop typing in notepad":
            self.listen_for_text = False
            self.update_chat_history("Stopped typing in Notepad.\n")
            self.read_out_loud("Stop typing.")
            return
        
        # ============for word document========================
        if query == "open word document":
            self.open_word_document()
            return
        
        if query == "start typing":
            if self.doc is not None:
                self.listen_for_text = True
                self.typing_mode = "word"
                self.read_out_loud("Start speaking.")
                self.read_out_loud("Please start speaking the text you want to add to the document.")
            return
        
        if query == "stop typing":
            self.listen_for_text = False  # New line to stop listening for text input
            if self.doc is not None:
                self.save_document()  # Save the document when user stops typing
                self.read_out_loud("Document saved.")
            return
        
        if query == "close word document":
            if self.doc is not None:
                self.close_word_document()
            else:
                self.update_chat_history("No document opened. Please open a Word document first.\n")
            return
        
        if query == "save word document":
            if self.doc is not None:
                self.save_document()
            else:
                self.update_chat_history("No document opened. Please open a Word document first.\n")
            return
        
        if query == "read aloud":
            if self.doc is not None:
                self.read_document_aloud()
            else:
                self.update_chat_history("No document opened. Please open a Word document first.\n")
            return
        
        
        # ==============================================================================================
        
        
    
        if self.listen_for_text:
            if query:

                # ================= WORD =================
                if self.typing_mode == "word" and self.doc is not None:
                    try:
                        if "\n" in query:
                            query = query.replace("\n", "")
                            self.doc.Content.InsertAfter("\n")

                        self.update_chat_history(f"User said: {query}\n")
                        self.doc.Content.InsertAfter(query + "\n")

                        self.read_out_loud("Text added.")

                    except Exception as e:
                        print("Word error:", e)
                        self.update_chat_history("Word connection lost.\n")
                        self.doc = None
                        self.typing_mode = None


            # ================= NOTEPAD =================
                elif self.typing_mode == "notepad" and self.notepad_open:
                    try:
                        import pygetwindow as gw
                        # Bring Notepad to front properly
                        windows = gw.getWindowsWithTitle("Notepad")
                        if windows:
                            windows[0].activate()
                            time.sleep(0.5)
                        if "\n" in query:
                            query = query.replace("\n", "")
                            pyautogui.press("enter")

                        self.update_chat_history(f"User said: {query}\n")
                        pyautogui.write(query + " ")

                        self.read_out_loud("Text added.")
                        return #this forces the chat not to move forward

                    except Exception as e:
                        print("Notepad error:", e)


        # ================= NOTHING =================
            else:
                # self.update_chat_history("No active typing mode.\n")

                return


        
        
        if query == "save document" and self.notepad_open:
            try:
                import pygetwindow as gw

                # Bring Notepad to front
                windows = gw.getWindowsWithTitle("Notepad")
                if windows:
                    windows[0].activate()
                    time.sleep(0.5)

                # Press Ctrl + S
                pyautogui.hotkey("ctrl", "s")
                time.sleep(1)

                # If it's first time save → type filename
                filename = f"note_{int(time.time())}.txt"
                pyautogui.write(filename)
                time.sleep(0.5)
                pyautogui.press("enter")

                self.update_chat_history(f"Document saved as {filename}\n")
                self.read_out_loud("Document saved.")

            except Exception as e:
                print("Save error:", e)
                self.update_chat_history("Failed to save document.\n")



        if query in ["where is document", "file location", "show file location"]:
            if self.notepad_file_path:
                self.update_chat_history(f"File is saved at:\n{self.notepad_file_path}\n")
                self.read_out_loud("Here is the file location.")
            else:
                self.update_chat_history("No saved file found.\n")
            return


    
        # ============for ss======================
        if query == "take screenshot":
            response = take_screenshot()
            self.update_chat_history(response)
            if "captured" in response:  
                pass
            return
        
        # =========for web search===========
        if query.startswith("search"):
            search_query = query[len("search"):].strip()
            self.update_chat_history(f"Searching for: {search_query}")
            webbrowser.open(f"https://www.google.com/search?q={search_query}")
            return

















        # ================== SPOTIFY DESKTOP CONTROL ==================

        if query.startswith("open spotify and play song"):
            song = query.replace("open spotify and play song", "").strip()

            try:
                os.startfile("spotify:")
                self.update_chat_history("Opening Spotify...")
                self.read_out_loud("Opening Spotify")
                time.sleep(5)

                #  Ensure window focus
                pyautogui.click(pyautogui.size()[0]//2, pyautogui.size()[1]//2)   # center of screen (adjust if needed)
                time.sleep(1)

                #  Open search (CORRECT SHORTCUT)
                pyautogui.hotkey("ctrl", "k")
                time.sleep(1)

                pyautogui.write(song)
                time.sleep(2)

                pyautogui.press("enter")
                time.sleep(2)

                #  Play first result
                pyautogui.press("tab", presses=2)
                time.sleep(1)
                pyautogui.press("enter")
            

                self.update_chat_history(f"Playing {song}")
                

            except Exception as e:
                self.update_chat_history(f"Error: {e}")
                self.read_out_loud("Failed to play song")

            return


        # ================== PLAY SONG DIRECT ==================

        if "play" in query:
            
            song = query.lower()
            # Remove possible keywords
            keywords = [
            "play song on spotify",
            "play on spotify",
            "play song",
            "play",
            "spotify"
                ]
            
            for word in keywords:
                song = song.replace(word, "")
                
            song = song.strip()
            
            try:
                os.startfile("spotify:")
                self.update_chat_history("Opening Spotify...")
                self.read_out_loud("Opening Spotify")
                time.sleep(5)

                #  Focus window
                pyautogui.click(960, 540)
                time.sleep(1)

                #  Correct search shortcut
                pyautogui.hotkey("ctrl", "k")
                time.sleep(1)

                pyautogui.write(song)
                time.sleep(2)

                pyautogui.press("enter")
                time.sleep(2)

                #  Play
                pyautogui.press("tab", presses=2)
                time.sleep(1)
                pyautogui.press("enter")
                time.sleep(1)
                pyautogui.press("space")

                self.update_chat_history(f"Playing {song}")
                self.read_out_loud(f"Playing {song}")

            except Exception as e:
                self.update_chat_history(f"Error: {e}")
                self.read_out_loud("Failed to play song")

            return
        



        # 🔹 INSTALL / DOWNLOAD APP
        if "install" in query or "download" in query:
            app_name = install_app_dynamic(query)

            if app_name:
                self.update_chat_history(f"Searching Microsoft Store for: {app_name}\n")
                self.read_out_loud(f"Searching for {app_name} in Microsoft Store")
            else:
                self.update_chat_history("Could not process app name\n")

            return
        
        
        
        
        
        # ================= CUSTOM CLOCK COMMANDS =================

        # ⏰ Alarm
        if "set alarm" in query:
            import re

            match = re.search(r'(\d{1,2})(?::(\d{2}))?\s*(am|pm)', query)

            if match:
                hour = int(match.group(1))
                minute = int(match.group(2)) if match.group(2) else 0
                period = match.group(3)

                self.set_alarm(hour, minute, period)
            else:
                    self.update_chat_history("Say like: set alarm for 7 am")
            return


        # ⏱ Timer
        if "set timer" in query:
            import re

            match = re.search(r'(\d+)', query)

            if match:
                value = int(match.group(1))

                if "minute" in query:
                    seconds = value * 60
                else:
                    seconds = value

                self.set_timer(seconds)
            else:
                self.update_chat_history("Say like: set timer for 5 minutes")
            return


        # ⏲ Stopwatch
        if "start stopwatch" in query:
            self.start_stopwatch()
            return

        if "stop stopwatch" in query:
            self.stop_stopwatch()
            return
        
        
        
        
        # ================= MEETING COMMAND =================

        if "join meeting" in query or "join this meeting" in query or "join meeting link" in query or "join zoom meeting" in query:
            self.join_meeting(query)
            return  
        
        
        # ====================for opening=======================
        if query.startswith("open"):
            app_name = query[len("open"):].strip().lower()
            if app_name == "camera":
                try:
                    subprocess.Popen(["start", "cmd", "/k", "start", "microsoft.windows.camera:"], shell=True)
                    self.camera_opened = True
                    self.update_chat_history(f"Opening {app_name}")
                except Exception as e:
                    self.update_chat_history(f"Error: {e}")
            # google is getting opened 
            elif app_name == "google":
                self.update_chat_history("Opening Google")
                webbrowser.open("https://www.google.com")    
            # youtube is getting opened 
            elif app_name == "youtube":
                self.update_chat_history("Opening YouTube")
                webbrowser.open("https://www.youtube.com")
            #notepad is getting opened  
            elif app_name == "notepad":
                subprocess.Popen(["notepad.exe"])
                self.notepad_open = True
                self.update_chat_history(f"Opening {app_name}")
            # calculator is getting opened
            elif app_name == "calculator":
                subprocess.Popen(["calc.exe"])
                self.update_chat_history(f"Opening {app_name}")
            # whatsapp is getting opened
            elif app_name == "whatsapp":
                os.startfile("whatsapp:")
                self.update_chat_history("Opening WhatsApp")
                return
            # spotify is getting opened
            elif app_name == "spotify":
                os.startfile("spotify:")
                self.update_chat_history("Opening Spotify")
                return
            # excel is getting opened
            elif app_name == "excel":
                os.startfile("excel.exe")
                self.update_chat_history(f"Opening {app_name}")
            # chrome is getting opened
            elif app_name == "chrome":
                os.startfile("chrome.exe")
                self.update_chat_history(f"Opening {app_name}")
            # zoom is getting opened
            elif app_name == "zoom":
                os.startfile("zoommtg:")
                self.update_chat_history("Opening Zoom")
                return
            # telegram is getting opened
            elif app_name == "telegram":
                try:
                    subprocess.Popen(["Telegram.exe"])
                except:
                    os.startfile("tg:")
                self.update_chat_history("Opening Telegram")
                return
            # paint is getting opened
            elif app_name == "paint":
                subprocess.Popen(["mspaint.exe"])
                self.update_chat_history("Opening Paint")
            # clock is getting opened
            elif app_name == "clock":
                os.startfile("ms-clock:")
                self.update_chat_history("Opening Clock")
                return
            # calendar is getting opened
            elif app_name == "outlook":
                os.startfile("outlookcal:")
                self.update_chat_history("Opening Outlook")
                return
            # settings is getting opened
            elif app_name == "settings":
                os.startfile("ms-settings:")
                self.update_chat_history("Opening Settings")
                return
            # microsoft edge is getting opened
            elif app_name == "edge":
                os.startfile("msedge.exe")
                self.update_chat_history("Opening Microsoft Edge")
            # microsoft store is getting opened
            elif app_name == "microsoft store":
                os.startfile("ms-windows-store:")
                self.update_chat_history(f"Opening Microsoft Store for {app_name}. Please click install.\n")
                self.read_out_loud(f"I found {app_name} in Microsoft Store. Please click install.")
                return
            # file explorer is getting opened
            elif app_name == "file explorer":
                subprocess.Popen(["explorer"])
                self.update_chat_history("Opening File Explorer")
            else:
                file_name = app_name.strip().lower()
                # 🔹 Convert spoken format to file format
                file_name = file_name.replace(" dot ", ".")
    
                # If user says "pdf", "docx" etc → convert to extension
                extensions = ["pdf", "docx", "txt", "xlsx", "pptx"]

                for ext in extensions:
                    if file_name.endswith(" " + ext):
                        file_name = file_name.replace(" " + ext, "." + ext)

                # Replace spaces with underscores (common in filenames)
                file_name = file_name.replace(" ", "_")

                # 🔹 Try opening file
                file_path = open_file_by_name(file_name)

                if file_path:
                    self.update_chat_history(f"Opening file: {file_path}\n")
                    self.read_out_loud("Opening your file")
                    return

                self.update_chat_history("File not found on this PC")
            return
        
        
        
        #========================= for closing===============
        if query.startswith("close"):
            close_app_name = query[len("close"):].strip().lower()
            
            # camera is getting closed
            if close_app_name == "camera":
                subprocess.Popen(["taskkill", "/F", "/IM", "WindowsCamera.exe"], shell=True)
                self.update_chat_history(f"closing {close_app_name}")
                return
            
            # these three apps are getting closed
            elif close_app_name == "google" or close_app_name == "youtube" or close_app_name =="whatsapp":
                subprocess.Popen(["taskkill", "/F", "/IM", "chrome.exe"], shell=True)
                self.update_chat_history("closing Google")
                return
            
            # youtube is getting closed
            elif close_app_name == "youtube":
                subprocess.Popen(["taskkill", "/F", "/IM", "chrome.exe"], shell=True)
                self.update_chat_history("closing YouTube")
                return
            
            # notepad is getting closed
            elif close_app_name == "notepad":
                subprocess.Popen(["taskkill", "/F", "/IM", "notepad.exe"], shell=True)
                self.update_chat_history("closing notepad")
                return
            
            # calculator is getting closed
            elif close_app_name == "calculator":
                subprocess.Popen(["taskkill", "/F", "/IM", "CalculatorApp.exe"], shell=True)
                self.update_chat_history("closing calculator")
                return
            
            # excel is getting closed
            elif close_app_name == "excel":
                subprocess.Popen(["taskkill", "/F", "/IM", "excel.exe"], shell=True)
                self.update_chat_history(f"closing {close_app_name}")
                return
            
            # spotify is getting closed
            elif close_app_name == "spotify":
                subprocess.Popen(["taskkill", "/F", "/IM", "spotify"], shell=True)
                self.update_chat_history(f"closing {close_app_name}")
                return
            
            # word document is getting closed
            elif close_app_name == "word document":
                subprocess.Popen(["taskkill", "/F", "/IM", "winword.exe"], shell=True)
                self.update_chat_history(f"closing {close_app_name}")
                return
            
            # telegram is getting closed
            elif close_app_name == "telegram":
                subprocess.Popen(["taskkill", "/F", "/IM", "Telegram.exe"], shell=True)
                self.update_chat_history(f"closing {close_app_name}")
                return
            
            # file explorer exceptions
            elif close_app_name == "file explorer":
                self.update_chat_history("Closing File Explorer is restricted for system stability")
                return
            
            # microsoft store is getting closed
            elif close_app_name == "microsoft store":
                subprocess.Popen(["taskkill", "/F", "/IM", "WinStore.App.exe"], shell=True)
                self.update_chat_history(f"closing {close_app_name}")
                return
            
            # outlook is not getting closed
            elif close_app_name == "outlook":
                subprocess.Popen(["taskkill", "/F", "/IM", "OUTLOOK.EXE"], shell=True)
                subprocess.Popen(["taskkill", "/F", "/IM", "olk.exe"], shell=True)
                self.update_chat_history("closing outlook")
                return
            
            # paint is getting closed
            elif close_app_name == "paint":
                subprocess.Popen(["taskkill", "/F", "/IM", "mspaint.exe"], shell=True)
                self.update_chat_history(f"closing {close_app_name}")
                return
            
            # settings is getting closed
            elif close_app_name == "settings":
                subprocess.Popen(["taskkill", "/F", "/IM", "SystemSettings.exe"], shell=True)
                self.update_chat_history(f"closing {close_app_name}")
                return
            
            # clock is getting closed
            elif close_app_name == "clock":
                subprocess.Popen(["taskkill", "/F", "/IM", "ClockApp.exe"], shell=True)
                self.update_chat_history(f"closing {close_app_name}")
                return

            
        # ==============FOR WHATSAPP======================
        if query.startswith("send whatsapp message"):
            try:
                # Example: send whatsapp message to rahul hello how are you
                parts = query.replace("send whatsapp message to", "").strip().split(" ", 1)
                
                if len(parts) < 2:
                    self.update_chat_history("Please say contact name and message")
                    return
                contact_name = parts[0]
                message = parts[1]

                # Open WhatsApp
                os.startfile("whatsapp:")
                self.update_chat_history("Opening WhatsApp...")
                time.sleep(6)

                # Open search using keyboard shortcut
                pyautogui.hotkey("ctrl", "f")
                time.sleep(1)

                # Type contact name
                pyautogui.write(contact_name)
                time.sleep(2)

                pyautogui.press("enter")
                time.sleep(1)

                # Type message
                pyautogui.write(message)
                time.sleep(1)

                pyautogui.press("enter")

                self.update_chat_history(f"Message sent to {contact_name}")
                self.read_out_loud("Message sent successfully")

            except Exception as e:
                self.update_chat_history(f"Error: {e}")
                self.read_out_loud("Failed to send message")
            return
        
        # ============FOR TELEGRAM=======================
        if query.startswith("send telegram message"):
            try:
                # Example: send whatsapp message to rahul hello how are you
                parts = query.replace("send telegram message to", "").strip().split(" ", 1)

                if len(parts) < 2:
                    self.update_chat_history("Please say contact name and message")
                    return
                contact_name = parts[0]
                message = parts[1]

                # Open Telegram
                os.startfile("tg:")
                self.update_chat_history("Opening Telegram...")
                time.sleep(6)

                # Open search using keyboard shortcut
                pyautogui.hotkey("ctrl", "f")
                time.sleep(1)

                # Type contact name
                pyautogui.write(contact_name)
                time.sleep(2)

                pyautogui.press("enter")
                time.sleep(1)

                # Type message
                pyautogui.write(message)
                time.sleep(1)

                pyautogui.press("enter")

                self.update_chat_history(f"Message sent to {contact_name}")
                self.read_out_loud("Message sent successfully")

            except Exception as e:
                self.update_chat_history(f"Error: {e}")
                self.read_out_loud("Failed to send message")
            return
        
        
        
        # ====================for youtube search=========================
        if query.startswith("youtube search"):
            search_query = query[len("youtube search"):].strip()
            self.update_chat_history(f"Searching YouTube for: {search_query}")
            webbrowser.open(f"https://www.youtube.com/results?search_query={search_query}")
            return
        
        
        
        
        # for date time & general one
        if "time" in query:
            current_time = datetime.now().strftime("%I:%M %p")
            self.update_chat_history(f"The current time is {current_time}.\n")
            self.read_out_loud(f"The current time is {current_time}.")
            return
        
        if "date" in query:
            current_date = datetime.now().strftime("%A, %B %d, %Y")
            self.update_chat_history(f"Today's date is {current_date}.\n")
            self.read_out_loud(f"Today's date is {current_date}.")
            return

        if query.startswith("what day is it"):
            current_day = datetime.now().strftime("%A")
            self.update_chat_history(f"Today is {current_day}.\n")
            self.read_out_loud(f"Today is {current_day}.")
            return

        try:
            chat_completion = self.client.chat.completions.create(
                messages=[
                    {"role": "system", "content": "You are Cypher, a helpful desktop AI assistant."},
                    {"role": "user", "content": query}
                ],
                 model="llama-3.1-8b-instant"
            )

            response_text = chat_completion.choices[0].message.content

        except Exception as e:
            response_text = f"Error: {str(e)}"

        self.update_chat_history(f"Assistant: {response_text}\n")
        self.read_out_loud(response_text)
        self.last_response = response_text

  
  
  

    def listen_voice_input(self):
        self.read_out_loud("Listening...")
        self.listening_label.setText("Listening...")
        recognizer = sr.Recognizer()
        with sr.Microphone() as source:
            print("Listening...")
            recognizer.adjust_for_ambient_noise(source)
            audio = recognizer.listen(source)

        try:
            query = recognizer.recognize_google(audio).lower()
            self.user_input.setText(query)

            # Check if the user asked to repeat
            if "repeat" in query or "what did you say" in query:
                if self.last_response:
                    self.update_chat_history(f"User: {query}\n")
                    self.update_chat_history(f"Assistant: {self.last_response}\n")
                    self.read_out_loud(self.last_response)
                    return
                else:
                    self.update_chat_history("Assistant: There is nothing to repeat.\n")
                    self.read_out_loud("There is nothing to repeat.")
                    return
            # Check for other commands
            elif query:
                self.send_message()
                
                
        except sr.UnknownValueError:
            print("Could not understand audio")
            self.update_chat_history("Could not understand audio. Please try again.\n")
            self.read_out_loud("Could not understand audio. Please try again.")    
            self.listening_label.setText("Could not understand audio")
        except sr.RequestError as e:
            print(f"Could not request results; {e}")
    
    

    def listen_for_repeat(self):
        recognizer = sr.Recognizer()
        with sr.Microphone() as source:
            recognizer.adjust_for_ambient_noise(source)
            audio = recognizer.listen(source)


    def update_chat_history(self, text):
        self.chat_history_display.append(text)  # Append the text to the response box
        self.chat_history_display.verticalScrollBar().setValue(self.chat_history_display.verticalScrollBar().maximum())



    def read_out_loud(self, text):
        if text and text.strip():
            self.tts_queue.put(text)


    
    
    def set_voice_gender(self, gender):
        if gender == "male":
            self.voice_gender = "male"
        elif gender == "female":
            self.voice_gender = "female"    

    def open_word_document(self):
        self.word_app, self.doc = open_word_document()
        self.update_chat_history("Word document opened.\n")
        self.read_out_loud("Word document opened.")

    def add_text_to_document(self):
        while self.listen_for_text:
            text = self.listen_voice_input()
            if text:
                if "stop typing" in text:
                    self.listen_for_text = False  # New line to stop listening for text input
                    self.update_chat_history("Document saved.\n")
                    self.save_document()
                    break
                
                if "\n" in text:
                    text = text.replace("\n", "")
                    self.doc.Content.InsertAfter("\n")
                
                self.update_chat_history(f"User said: {text}\n")
                self.doc.Content.InsertAfter(text + "\n")
                
                self.read_out_loud("Text added. Keep typing or say 'stop typing' to finish.")
            else:
                self.read_out_loud("No text detected. Please try again.")
    
    def save_document(self):
        filename = save_document(self.doc)
        self.update_chat_history(f"Document saved as '{filename}'.\n")
        
    
    
    
    def close_word_document(self):
        close_word_document(self.word_app, self.doc)
        self.doc = None
        self.word_app = None
        self.typing_mode = None
        self.update_chat_history("Word document closed.\n")
        self.read_out_loud("Word document closed.")

    def read_document_aloud(self):
        try:
            if self.doc:
                content = self.doc.Content.Text.strip()

                if content:
                    self.update_chat_history("Reading document...\n")
                    self.read_out_loud(content[:2000])  # 🔹 limit (important)
                else:
                    self.update_chat_history("Document is empty.\n")
            else:
                self.update_chat_history("No document opened.\n")

        except Exception as e:
            print("Read error:", e)
            
            
    
    
    
    # ================= CUSTOM CLOCK SYSTEM =================

    def set_alarm(self, hour, minute, period):
        def alarm_thread():
            self.update_chat_history(f"Alarm set for {hour}:{minute:02d} {period}")
            self.read_out_loud(f"Alarm set for {hour} {minute} {period}")

            while True:
                now = datetime.now()

                if period == "pm" and hour != 12:
                    hour_24 = hour + 12
                elif period == "am" and hour == 12:
                    hour_24 = 0
                else:
                    hour_24 = hour

                if now.hour == hour_24 and now.minute == minute:
                    self.update_chat_history("⏰ Alarm ringing!")
                    self.read_out_loud("Wake up! Alarm ringing!")

                    for _ in range(5):
                        winsound.Beep(1000, 500)  # Beep at 1000 Hz for 500 ms  
                        time.sleep(1)
                    break

                time.sleep(10)

        threading.Thread(target=alarm_thread, daemon=True).start()


    def set_timer(self, seconds):
        def timer_thread():
            self.update_chat_history(f"Timer set for {seconds} seconds")
            self.read_out_loud(f"Timer set for {seconds} seconds")

            time.sleep(seconds)

            self.update_chat_history("⏰ Timer finished!")
            self.read_out_loud("Time is up!")


            for _ in range(5):
                winsound.Beep(1000, 500)  # Beep at 1000 Hz for 500 ms  
                

        threading.Thread(target=timer_thread, daemon=True).start()


    def start_stopwatch(self):
        self.stopwatch_running = True
        self.stopwatch_start = time.time()
        
        # 🔊 Beep when starting
        winsound.Beep(1000, 300)


        def run():
            while self.stopwatch_running:
                elapsed = int(time.time() - self.stopwatch_start)
                print(f"Stopwatch: {elapsed} sec", end="\r")
                time.sleep(1)

        threading.Thread(target=run, daemon=True).start()

        self.update_chat_history("Stopwatch started")
        self.read_out_loud("Stopwatch started")


    def stop_stopwatch(self):
        if hasattr(self, "stopwatch_running") and self.stopwatch_running:
            self.stopwatch_running = False
            elapsed = int(time.time() - self.stopwatch_start)
            # 🔊 Beep when starting
            winsound.Beep(800, 200)
            time.sleep(0.5)
            winsound.Beep(800, 200)
            


            self.update_chat_history(f"Stopwatch stopped at {elapsed} seconds")
            self.read_out_loud(f"Stopwatch stopped at {elapsed} seconds")
        else:
            self.update_chat_history("Stopwatch is not running")
            
            
            
            
    # ================= MEETING JOIN SYSTEM =================

    def join_meeting(self, query):
        import re

        self.update_chat_history("Joining meeting...")
        self.read_out_loud("Joining your meeting")

        # 🔹 Extract link or ID
        link_match = re.search(r'(https?://\S+)', query)

        if link_match:
            link = link_match.group(1)
            webbrowser.open(link)
            time.sleep(8)

            # 🔹 Try auto join
            pyautogui.press("tab", presses=5)
            pyautogui.press("enter")

            return

        # 🔹 Zoom ID case
        id_match = re.search(r'(\d{9,11})', query)

        if id_match:
            meeting_id = id_match.group(1)

            # Open Zoom with meeting ID
            os.startfile(f"zoommtg://zoom.us/join?confno={meeting_id}")
            time.sleep(8)

            pyautogui.press("enter")
            return

        self.update_chat_history("No valid meeting link or ID found")
            
            
    def display_help(self):
        if os.path.exists(self.manual_path):
            try:
                webbrowser.open(self.manual_path)
            except Exception as e:
                self.update_chat_history(f"Error opening user manual: {e}\n")
        else:
            self.update_chat_history("User manual not found.\n")
    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ChatWindow()
    sys.exit(app.exec_())
