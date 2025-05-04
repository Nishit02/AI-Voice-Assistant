import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import smtplib
from collections import deque
import heapq
import itertools  # For TSP permutation approach
import json  # For saving/loading TSP data
import subprocess  # For system control (platform-dependent)
import platform  # For OS detection
from PIL import ImageGrab  # For taking screenshots (Windows)
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume  # For volume control (Windows)
from comtypes import CLSCTX_ALL
from ctypes import cast, POINTER  # Import missing modules
import requests  # For weather and web searches
import psutil  # For system monitoring
import pyautogui  # For additional system control
import wolframalpha  # For calculations and knowledge
import pyjokes  # For jokes
import speedtest  # For internet speed test
from bs4 import BeautifulSoup  # For web scraping
import calendar
import sys
import time
from googlesearch import search  # For Google search results
import screen_brightness_control as sbc  # For brightness control
from win32gui import GetWindowText, GetForegroundWindow  # For window management
import spotipy  # For Spotify control
from spotipy.oauth2 import SpotifyOAuth
import winreg  # For getting installed apps
import win32com.client  # For creating shortcuts
import math  # For calculations

# Initialize TTS engine
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)


def speak(audio):
    print("Jarvis:", audio)
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour < 12:
        speak("Good Morning!")
    elif hour < 18:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")
    speak("I am Jarvis. How can I assist you today?")


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        try:
            audio = r.listen(source)
        except sr.WaitTimeoutError:
            print("No speech detected.")
            return "None"
        except Exception as e:
            print(f"Error during listening: {e}")
            return "None"

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in').lower()
        print(f"You said: {query}")
        return query
    except sr.UnknownValueError:
        speak("I could not understand that. Please say that again.")
        return "None"
    except sr.RequestError as e:
        speak(f"Could not request results from Google Speech Recognition service; {e}")
        return "None"


def sendEmail(to, content):
    # --- SECURITY WARNING: DO NOT STORE YOUR PASSWORD IN CODE! ---
    # Consider using environment variables or a secure configuration method.
    sender_email = "nishitlight@gmail.com"  # Replace with your email
    sender_password = "YOUR_PASSWORD"  # REPLACE WITH YOUR ACTUAL PASSWORD

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, to, content)
        server.quit()
        speak("Email has been sent!")
    except Exception as e:
        print(f"Error sending email: {e}")
        speak("Sorry, I can't send the email right now.")


# --- Traveling Salesman Problem Implementation ---
def solve_tsp(distances):
    """
    Solves the Traveling Salesman Problem using a brute-force approach (permutations).
    This is suitable for a small number of cities.
    """
    num_cities = len(distances)
    if num_cities == 0:
        return [], 0
    start_city = 0
    cities = list(range(num_cities))
    cities.remove(start_city)

    shortest_route = None
    min_distance = float('inf')

    for path in itertools.permutations(cities):
        current_path = [start_city] + list(path) + [start_city]
        current_distance = 0
        for i in range(num_cities):
            current_distance += distances[current_path[i]][current_path[i + 1]]

        if current_distance < min_distance:
            min_distance = current_distance
            shortest_route = current_path

    return shortest_route, min_distance


def get_tsp_input():
    """
    Gets the number of cities via voice or text and distances between them via voice or text.
    """
    speak("Do you want to provide the number of cities by voice or by typing?")
    method_choice_cities = takeCommand().lower()

    if 'voice' in method_choice_cities:
        speak("Please say the number of cities.")
        while True:
            num_cities_str = takeCommand()
            try:
                num_cities = int(num_cities_str)
                if num_cities > 0:
                    break
                else:
                    speak("Number of cities must be greater than 0. Please try again.")
            except (ValueError, TypeError):
                speak("Sorry, I didn't understand that. Please say the number of cities again.")
    elif 'typing' in method_choice_cities:
        while True:
            try:
                num_cities_str = input("Enter the number of cities: ")
                num_cities = int(num_cities_str)
                if num_cities > 0:
                    speak(f"Number of cities set to {num_cities}")
                    break
                else:
                    print("Number of cities must be greater than 0. Please try again.")
                    speak("Number of cities must be greater than 0. Please try again.")
            except ValueError:
                print("Invalid input. Please enter a number.")
                speak("Invalid input. Please enter a number.")
    else:
        speak("Invalid choice. Using voice input for the number of cities.")
        speak("Please say the number of cities.")
        while True:
            num_cities_str = takeCommand()
            try:
                num_cities = int(num_cities_str)
                if num_cities > 0:
                    break
                else:
                    speak("Number of cities must be greater than 0. Please try again.")
            except (ValueError, TypeError):
                speak("Sorry, I didn't understand that. Please say the number of cities again.")

    distances = [[0.0 for _ in range(num_cities)] for _ in range(num_cities)]
    speak("Now, let's enter the distances between the cities.")

    word_to_num = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
                    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10}

    for i in range(num_cities):
        for j in range(i + 1, num_cities):
            speak(f"Do you want to provide the distance between city {i + 1} and city {j + 1} by voice or by typing?")
            method_choice_distance = takeCommand().lower()

            while True:
                if 'voice' in method_choice_distance:
                    speak(f"Please say the distance between city {i + 1} and city {j + 1} in the format 'distance is Z'.")
                    distance_query = takeCommand()
                    if distance_query:
                        try:
                            parts = distance_query.lower().split()
                            if 'distance' in parts and 'is' in parts:
                                try:
                                    distance_index = parts.index('is') + 1
                                    distance_val_str = parts[distance_index]
                                    distance = float(distance_val_str)
                                    distances[i][j] = distance
                                    distances[j][i] = distance
                                    speak(f"Distance between city {i + 1} and city {j + 1} is set to {distance}")
                                    break
                                except (ValueError, IndexError):
                                    speak("Invalid format. Please use 'distance is Z'.")
                            else:
                                speak("Invalid format. Please use 'distance is Z'.")
                        except Exception as e:
                            print(f"Error processing distance input: {e}")
                            speak("Sorry, I didn't understand that. Please try again.")
                    else:
                        speak("No input received. Please try again.")

                elif 'typing' in method_choice_distance:
                    while True:
                        try:
                            distance_str = input(f"Enter the distance between city {i + 1} and city {j + 1}: ")
                            distance = float(distance_str)
                            distances[i][j] = distance
                            distances[j][i] = distance
                            speak(f"Distance between city {i + 1} and city {j + 1} is set to {distance}")
                            break
                        except ValueError:
                            print("Invalid input. Please enter a number.")
                            speak("Invalid input. Please enter a number.")
                    break  # Break out of the inner while loop for typing

                else:
                    speak("Invalid choice. Please say 'voice' or 'typing'.")
                    method_choice_distance = takeCommand().lower()

    return distances


def save_tsp_data(distances, filename="tsp_data.json"):
    try:
        with open(filename, 'w') as f:
            json.dump(distances, f)
        speak(f"TSP data saved to {filename}")
    except Exception as e:
        print(f"Error saving TSP data: {e}")
        speak("Could not save TSP data.")


def load_tsp_data(filename="tsp_data.json"):
    try:
        with open(filename, 'r') as f:
            data = json.load(f)
        speak(f"TSP data loaded from {filename}")
        return data
    except FileNotFoundError:
        speak("TSP data file not found.")
        return None
    except Exception as e:
        print(f"Error loading TSP data: {e}")
        speak("Could not load TSP data.")
        return None


def control_volume(action="up"):
    os_name = platform.system().lower()
    try:
        if "windows" in os_name:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            current_volume = volume.GetMasterVolumeLevelScalar()

            if action == "up":
                new_volume = min(1.0, current_volume + 0.1)
            elif action == "down":
                new_volume = max(0.0, current_volume - 0.1)
            elif action == "mute":
                volume.SetMute(1, None)
                speak("Volume muted.")
                return
            elif action == "unmute":
                volume.SetMute(0, None)
                speak("Volume unmuted.")
                return
            else:
                speak("Invalid volume action.")
                return

            volume.SetMasterVolumeLevelScalar(new_volume, None)
            speak(f"Volume set to {int(new_volume * 100)} percent.")

        elif "linux" in os_name:
            if action == "up":
                subprocess.run(["amixer", "set", "Master", "5%+"], capture_output=True)
            elif action == "down":
                subprocess.run(["amixer", "set", "Master", "5%-"], capture_output=True)
            elif action == "mute":
                subprocess.run(["amixer", "set", "Master", "mute"], capture_output=True)
                speak("Volume muted.")
                return
            elif action == "unmute":
                subprocess.run(["amixer", "set", "Master", "unmute"], capture_output=True)
                return
            else:
                speak("Invalid volume action.")
                return
            current_volume_result = subprocess.run(["amixer", "get", "Master"], capture_output=True, text=True)
            if "[on]" in current_volume_result.stdout:
                volume_level = current_volume_result.stdout.split("%")[0].split("[")[-1]
                speak(f"Volume set to approximately {volume_level} percent.")

        elif "darwin" in os_name:  # macOS
            if action == "up":
                subprocess.run(["osascript", "-e", "set volume output volume (output volume of (get volume settings)) + 10"], capture_output=True)
            elif action == "down":
                subprocess.run(["osascript", "-e", "set volume output volume (output volume of (get volume settings)) - 10"], capture_output=True)
            elif action == "mute":
                subprocess.run(["osascript", "-e", "set volume output muted true"], capture_output=True)
                speak("Volume muted.")
                return
            elif action == "unmute":
                subprocess.run(["osascript", "-e", "set volume output muted false"], capture_output=True)
                return
            else:
                speak("Invalid volume action.")
                return
            volume_result = subprocess.run(["osascript", "-e", "output volume of (get volume settings)"], capture_output=True, text=True)
            try:
                volume_level = int(volume_result.stdout.strip())
                speak(f"Volume set to approximately {volume_level} percent.")
            except ValueError:
                pass

        else:
            speak("Volume control is not supported on this operating system.")

    except Exception as e:
        print(f"Error controlling volume: {e}")
        speak("Could not control volume.")


def take_screenshot(filename="screenshot.png"):
    os_name = platform.system().lower()
    try:
        if "windows" in os_name:
            screenshot = ImageGrab.grab()
            screenshot.save(filename)
            speak(f"Screenshot saved as {filename}")
        elif "linux" in os_name:
            subprocess.run(["scrot", filename], check=True)
            speak(f"Screenshot saved as {filename}")
        elif "darwin" in os_name:  # macOS
            subprocess.run(["screencapture", filename], check=True)
            speak(f"Screenshot saved as {filename}")
        else:
            speak("Screenshot functionality is not supported on this operating system.")
    except FileNotFoundError:
        speak("Screenshot tool not found. Please install it (e.g., 'scrot' on Linux).")
    except Exception as e:
        print(f"Error taking screenshot: {e}")
        speak("Could not take screenshot.")


def lock_computer():
    os_name = platform.system().lower()
    try:
        if "windows" in os_name:
            subprocess.run(["rundll32.exe", "user32.dll,LockWorkStation"])
            speak("Computer locked.")
        elif "linux" in os_name:
            subprocess.run(["gnome-screensaver-command", "--lock"])  # Might need different command for other DEs
            speak("Computer locked.")
        elif "darwin" in os_name:  # macOS
            subprocess.run(["/System/Library/CoreServices/Menu Extras/User.menu/Contents/Resources/CGSession", "-suspend"])
            speak("Computer locked.")
        else:
            speak("Lock computer functionality is not supported on this operating system.")
    except FileNotFoundError:
        speak("Lock screen tool not found.")
    except Exception as e:
        print(f"Error locking computer: {e}")
        speak("Could not lock computer.")


def shutdown_computer():
    speak("Initiating shutdown sequence. Are you sure?")
    confirmation = takeCommand()
    if "shutdown" in confirmation:
        os_name = platform.system().lower()
        try:
            if "windows" in os_name:
                subprocess.run(["shutdown", "/s", "/t", "1"])
            elif "linux" in os_name:
                subprocess.run(["sudo", "shutdown", "now"])  # Requires sudo passwordless or manual entry
                speak("Shutting down.")
            elif "darwin" in os_name:  # macOS
                subprocess.run(["sudo", "shutdown", "-h", "now"])  # Requires sudo passwordless or manual entry
                speak("Shutting down.")
            else:
                speak("Shutdown functionality is not supported on this operating system.")
        except Exception as e:
            print(f"Error shutting down: {e}")
            speak("Could not initiate shutdown.")
    else:
        speak("Shutdown cancelled.")


def restart_computer():
    speak("Initiating restart sequence. Are you sure?")
    confirmation = takeCommand()
    if "yes" in confirmation:
        os_name = platform.system().lower()
        try:
            if "windows" in os_name:
                subprocess.run(["shutdown", "/r", "/t", "1"])
            elif "linux" in os_name:
                subprocess.run(["sudo", "reboot"])  # Requires sudo passwordless or manual entry
                speak("Restarting.")
            elif "darwin" in os_name:  # macOS
                subprocess.run(["sudo", "reboot"])  # Requires sudo passwordless or manual entry
                speak("Restarting.")
            else:
                speak("Restart functionality is not supported on this operating system.")
        except Exception as e:
            print(f"Error restarting: {e}")
            speak("Could not initiate restart.")
    else:
        speak("Restart cancelled.")


def get_weather(city="Chennai"):
    """Get weather information for a city"""
    api_key = "d850f7f52bf19300a9eb4b0aa6b80f0d"  # OpenWeather API key
    base_url = "http://api.openweathermap.org/data/2.5/weather"
    
    try:
        params = {
            "q": city,
            "appid": api_key,
            "units": "metric"
        }
        response = requests.get(base_url, params=params)
        data = response.json()
        
        if response.status_code == 200:
            temp = data["main"]["temp"]
            humidity = data["main"]["humidity"]
            desc = data["weather"][0]["description"]
            feels_like = data["main"]["feels_like"]
            wind_speed = data["wind"]["speed"]
            speak(f"The temperature in {city} is {temp}°C, feels like {feels_like}°C, with {humidity}% humidity. Wind speed is {wind_speed} meters per second. Weather condition: {desc}")
        else:
            speak("Sorry, I couldn't fetch the weather information.")
    except Exception as e:
        print(f"Error getting weather: {e}")
        speak("Sorry, I couldn't fetch the weather information.")


def unlock_computer():
    """Unlock the computer using different methods based on OS"""
    os_name = platform.system().lower()
    try:
        if "windows" in os_name:
            # On Windows, we can simulate Ctrl+Alt+Del followed by Enter
            pyautogui.hotkey('ctrl', 'alt', 'del')
            time.sleep(1)  # Wait for the menu to appear
            pyautogui.press('enter')  # Press Enter to unlock
            speak("Attempting to unlock computer.")
        elif "linux" in os_name:
            # For Linux, depends on desktop environment
            subprocess.run(["loginctl", "unlock-session"])
            speak("Attempting to unlock computer.")
        elif "darwin" in os_name:  # macOS
            # For macOS, simulating keyboard input
            pyautogui.press('space')
            time.sleep(1)
            pyautogui.press('enter')
            speak("Attempting to unlock computer.")
        else:
            speak("Unlock computer functionality is not supported on this operating system.")
    except Exception as e:
        print(f"Error unlocking computer: {e}")
        speak("Could not unlock computer.")


def google_search(query):
    """Perform a Google search and read out the top results"""
    try:
        speak(f"Searching Google for {query}")
        # First, directly open the search in browser
        search_url = f"https://www.google.com/search?q={query}"
        webbrowser.open(search_url)
        speak("I've opened the search results in your browser.")
        
        # Then get and speak the results
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(search_url, headers=headers)
        
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            search_results = []
            
            # Find search result divs
            for result in soup.find_all('div', class_='g')[:5]:
                title_element = result.find('h3')
                link_element = result.find('a')
                
                if title_element and link_element:
                    title = title_element.text
                    link = link_element.get('href')
                    if link and link.startswith('http'):
                        search_results.append((title, link))
            
            if search_results:
                speak("Here are the top results I found:")
                for idx, (title, link) in enumerate(search_results, 1):
                    speak(f"Result {idx}: {title}")
                    print(f"Link {idx}: {link}")
            else:
                speak("I couldn't extract the search results, but I've opened the search page in your browser.")
        else:
            speak("I've opened the search in your browser, but couldn't fetch additional details.")
    except Exception as e:
        print(f"Error during Google search: {e}")
        speak("I encountered an error, but I'll try to open the search in your browser.")
        try:
            webbrowser.open(f"https://www.google.com/search?q={query}")
        except:
            speak("Sorry, I couldn't perform the search operation.")


def get_system_info():
    """Get system resource information"""
    cpu_percent = psutil.cpu_percent()
    memory = psutil.virtual_memory()
    disk = psutil.disk_usage('/')
    
    speak(f"CPU usage is {cpu_percent}%")
    speak(f"Memory usage is {memory.percent}%")
    speak(f"Disk usage is {disk.percent}%")


def tell_joke():
    """Tell a random joke"""
    joke = pyjokes.get_joke()
    speak(joke)


def check_internet_speed():
    """Check internet connection speed"""
    speak("Testing internet speed. This might take a moment...")
    st = speedtest.Speedtest()
    
    download_speed = st.download() / 1_000_000  # Convert to Mbps
    upload_speed = st.upload() / 1_000_000  # Convert to Mbps
    
    speak(f"Download speed is {download_speed:.2f} Mbps")
    speak(f"Upload speed is {upload_speed:.2f} Mbps")


def calculate_expression(query):
    """Use WolframAlpha to calculate mathematical expressions"""
    app_id = "UVJR3T-8JG6WGJ374"  # WolframAlpha App ID
    client = wolframalpha.Client(app_id)
    
    try:
        res = client.query(query)
        answer = next(res.results).text
        speak(f"The answer is {answer}")
    except Exception as e:
        print(f"Error during calculation: {e}")
        speak("Sorry, I couldn't calculate that.")


def get_calendar_info():
    """Get calendar information for current month"""
    now = datetime.datetime.now()
    cal = calendar.month(now.year, now.month)
    speak(f"Here's the calendar for {calendar.month_name[now.month]} {now.year}")
    print(cal)


def take_note(text):
    """Save a note to a text file"""
    date = datetime.datetime.now()
    file_name = f"note_{date.strftime('%Y%m%d_%H%M%S')}.txt"
    with open(file_name, "w") as f:
        f.write(text)
    speak(f"I've made a note of that and saved it as {file_name}")


def open_whatsapp_chat(contact_name):
    """Open WhatsApp chat with specific contact"""
    try:
        # Open WhatsApp Web
        webbrowser.open("https://web.whatsapp.com/")
        speak("Opening WhatsApp Web. Please wait for it to load...")
        time.sleep(15)  # Wait for WhatsApp Web to load
        
        # Click on search box
        pyautogui.hotkey('ctrl', 'f')
        time.sleep(1)
        
        # Type contact name
        pyautogui.write(contact_name)
        time.sleep(2)
        
        # Press enter to select the contact
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.press('escape')  # Clear search
        speak(f"Opened chat with {contact_name}")
    except Exception as e:
        print(f"Error opening WhatsApp chat: {e}")
        speak("Sorry, I couldn't open the WhatsApp chat.")


def send_whatsapp_message(contact_name, message):
    """Send WhatsApp message to specific contact"""
    try:
        open_whatsapp_chat(contact_name)
        time.sleep(2)
        
        # Type and send message
        pyautogui.write(message)
        time.sleep(1)
        pyautogui.press('enter')
        speak(f"Message sent to {contact_name}")
    except Exception as e:
        print(f"Error sending WhatsApp message: {e}")
        speak("Sorry, I couldn't send the WhatsApp message.")


def whatsapp_call(contact_name, video=False):
    """Make WhatsApp call (voice or video) to specific contact"""
    try:
        open_whatsapp_chat(contact_name)
        time.sleep(2)
        
        # Click call button
        if video:
            pyautogui.click(pyautogui.locateCenterOnScreen('video_call_button.png'))
            speak(f"Starting video call with {contact_name}")
        else:
            pyautogui.click(pyautogui.locateCenterOnScreen('voice_call_button.png'))
            speak(f"Starting voice call with {contact_name}")
    except Exception as e:
        print(f"Error making WhatsApp call: {e}")
        speak("Sorry, I couldn't make the WhatsApp call.")


def control_brightness(action="up"):
    """Control screen brightness"""
    try:
        current = sbc.get_brightness()[0]
        if action == "up":
            new_brightness = min(100, current + 10)
            sbc.set_brightness(new_brightness)
            speak(f"Brightness increased to {new_brightness} percent")
        elif action == "down":
            new_brightness = max(0, current - 10)
            sbc.set_brightness(new_brightness)
            speak(f"Brightness decreased to {new_brightness} percent")
        elif action == "set":
            speak("What percentage of brightness would you like?")
            try:
                percent = int(takeCommand())
                if 0 <= percent <= 100:
                    sbc.set_brightness(percent)
                    speak(f"Brightness set to {percent} percent")
                else:
                    speak("Brightness percentage must be between 0 and 100")
            except:
                speak("Sorry, I couldn't understand the brightness level")
    except Exception as e:
        print(f"Error controlling brightness: {e}")
        speak("Sorry, I couldn't control the brightness")


def setup_spotify():
    """Setup Spotify API client"""
    try:
        # Your Spotify API credentials
        client_id = "843950f76a214a93a8de0d8f85d813da"     # Your Spotify Client ID
        client_secret = "eadc42a98f094d94b6c5070cd4c0dad9"  # Your Spotify Client Secret
        redirect_uri = "http://127.0.0.1:8888/callback"     # Using localhost IP address directly
        
        # Define the scope of permissions
        scope = "user-read-playback-state user-modify-playback-state user-read-currently-playing playlist-read-private"
        
        # Create the SpotifyOAuth manager with cache handling
        auth_manager = SpotifyOAuth(
            client_id=client_id,
            client_secret=client_secret,
            redirect_uri=redirect_uri,
            scope=scope,
            open_browser=True,
            cache_path=".spotify_cache",  # Save tokens in a cache file
            requests_timeout=10,
            requests_session=True
        )
        
        # Create and return the Spotify client
        sp = spotipy.Spotify(auth_manager=auth_manager)
        
        # Test the connection
        try:
            sp.current_user()  # Test if we're properly authenticated
            speak("Successfully connected to Spotify")
            return sp
        except Exception as auth_error:
            print(f"Authentication error: {auth_error}")
            speak("Failed to authenticate with Spotify. Please check your credentials and try again.")
            return None
            
    except Exception as e:
        print(f"Error setting up Spotify: {e}")
        speak("There was an error setting up Spotify. Please check your credentials and redirect URI.")
        return None


def control_spotify(action, sp=None):
    """Control Spotify playback"""
    if sp is None:
        sp = setup_spotify()
        if sp is None:
            speak("Could not connect to Spotify. Please check your credentials and internet connection.")
            return
    
    try:
        if action == "play":
            sp.start_playback()
            speak("Playing Spotify")
        elif action == "pause":
            sp.pause_playback()
            speak("Paused Spotify")
        elif action == "next":
            sp.next_track()
            speak("Playing next track")
        elif action == "previous":
            sp.previous_track()
            speak("Playing previous track")
        elif action == "current":
            current = sp.current_playback()
            if current and current.get('item'):
                track_name = current['item']['name']
                artist_name = current['item']['artists'][0]['name']
                speak(f"Currently playing {track_name} by {artist_name}")
            else:
                speak("No track is currently playing")
    except Exception as e:
        print(f"Error controlling Spotify: {e}")
        speak("Sorry, I couldn't control Spotify. Make sure Spotify is running and you're logged in.")


def get_installed_apps():
    """Get list of installed applications"""
    apps = []
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
            for i in range(winreg.QueryInfoKey(key)[0]):
                try:
                    subkey_name = winreg.EnumKey(key, i)
                    with winreg.OpenKey(key, subkey_name) as subkey:
                        try:
                            display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                            apps.append(display_name)
                        except:
                            continue
                except:
                    continue
    except Exception as e:
        print(f"Error getting installed apps: {e}")
    return apps


def open_application(app_name):
    """Open a specified application"""
    try:
        # Common paths for popular applications
        app_paths = {
            "spotify": r"C:\Users\{username}\AppData\Roaming\Spotify\Spotify.exe",
            "chrome": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            "firefox": r"C:\Program Files\Mozilla Firefox\firefox.exe",
            "word": r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
            "excel": r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
            "powerpoint": r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
            "notepad": r"C:\Windows\System32\notepad.exe",
            "calculator": r"C:\Windows\System32\calc.exe"
        }
        
        app_name = app_name.lower()
        username = os.getenv("USERNAME")
        
        if app_name in app_paths:
            path = app_paths[app_name].format(username=username)
            if os.path.exists(path):
                os.startfile(path)
                speak(f"Opening {app_name}")
                return
        
        # Try using Run command
        try:
            os.system(f"start {app_name}")
            speak(f"Attempting to open {app_name}")
        except:
            speak(f"Sorry, I couldn't find {app_name}")
    except Exception as e:
        print(f"Error opening application: {e}")
        speak(f"Sorry, I couldn't open {app_name}")


def close_application(app_name):
    """Close a specified application"""
    try:
        os.system(f"taskkill /f /im {app_name}.exe")
        speak(f"Closed {app_name}")
    except Exception as e:
        print(f"Error closing application: {e}")
        speak(f"Sorry, I couldn't close {app_name}")


def solve_missionaries_cannibals():
    """
    Solve the Missionaries and Cannibals problem using BFS.
    User can input the number of missionaries and cannibals.
    At no point should cannibals outnumber missionaries on either bank.
    """
    def get_number_input(prompt, input_type="voice"):
        while True:
            if input_type == "voice":
                speak(prompt)
                response = takeCommand().lower()
                try:
                    # Convert word numbers to digits
                    word_to_num = {
                        'one': '1', 'two': '2', 'three': '3', 'four': '4', 'five': '5',
                        'six': '6', 'seven': '7', 'eight': '8', 'nine': '9', 'ten': '10'
                    }
                    for word, num in word_to_num.items():
                        response = response.replace(word, num)
                    number = int(''.join(filter(str.isdigit, response)))
                    if number > 0:
                        return number
                    else:
                        speak("Please provide a positive number.")
                except:
                    speak("I couldn't understand that number. Please try again.")
            else:  # typing
                try:
                    number = int(input(prompt))
                    if number > 0:
                        return number
                    else:
                        print("Please provide a positive number.")
                except ValueError:
                    print("Invalid input. Please enter a number.")

    # Get input method preference
    speak("Would you like to provide input by voice or typing?")
    input_method = takeCommand().lower()
    input_type = "voice" if "voice" in input_method else "typing"

    # Get number of missionaries and cannibals
    if input_type == "voice":
        speak("Please say the number of missionaries")
        num_missionaries = get_number_input("How many missionaries?", "voice")
        speak(f"Got it, {num_missionaries} missionaries.")
        
        speak("Now, please say the number of cannibals")
        num_cannibals = get_number_input("How many cannibals?", "voice")
        speak(f"Got it, {num_cannibals} cannibals.")
    else:
        num_missionaries = get_number_input("Enter the number of missionaries: ", "typing")
        print(f"Number of missionaries: {num_missionaries}")
        
        num_cannibals = get_number_input("Enter the number of cannibals: ", "typing")
        print(f"Number of cannibals: {num_cannibals}")

    # Get boat capacity
    if input_type == "voice":
        speak("What is the boat capacity?")
        boat_capacity = get_number_input("How many people can the boat carry?", "voice")
    else:
        boat_capacity = get_number_input("Enter the boat capacity: ", "typing")
    
    class State:
        def __init__(self, m_left, c_left, boat_left, m_right, c_right, parent=None, move=None):
            self.m_left = m_left
            self.c_left = c_left
            self.boat_left = boat_left
            self.m_right = m_right
            self.c_right = c_right
            self.parent = parent
            self.move = move

        def is_valid(self):
            # Check if numbers are valid
            if (self.m_left < 0 or self.c_left < 0 or 
                self.m_right < 0 or self.c_right < 0):
                return False
            # Check if missionaries are outnumbered on either bank
            if (self.m_left > 0 and self.m_left < self.c_left) or \
               (self.m_right > 0 and self.m_right < self.c_right):
                return False
            return True

        def is_goal(self):
            # Goal state: all missionaries and cannibals on right bank
            return self.m_left == 0 and self.c_left == 0 and not self.boat_left

        def __eq__(self, other):
            return (self.m_left == other.m_left and 
                    self.c_left == other.c_left and 
                    self.boat_left == other.boat_left)

        def __hash__(self):
            return hash((self.m_left, self.c_left, self.boat_left))

    def get_next_states(state):
        states = []
        # Generate all possible moves within boat capacity
        moves = []
        for m in range(boat_capacity + 1):
            for c in range(boat_capacity + 1):
                if 1 <= m + c <= boat_capacity:
                    moves.append((m, c))
        
        # Current bank is left if boat is on left, otherwise right
        if state.boat_left:
            m_current = state.m_left
            c_current = state.c_left
        else:
            m_current = state.m_right
            c_current = state.c_right

        for m, c in moves:
            if m_current >= m and c_current >= c:  # Check if move is possible
                if state.boat_left:
                    new_state = State(
                        state.m_left - m,
                        state.c_left - c,
                        False,
                        state.m_right + m,
                        state.c_right + c,
                        state,
                        f"Move {m} missionaries and {c} cannibals to right bank"
                    )
                else:
                    new_state = State(
                        state.m_left + m,
                        state.c_left + c,
                        True,
                        state.m_right - m,
                        state.c_right - c,
                        state,
                        f"Move {m} missionaries and {c} cannibals to left bank"
                    )
                if new_state.is_valid():
                    states.append(new_state)
        return states

    def solve():
        initial_state = State(num_missionaries, num_cannibals, True, 0, 0)
        if not initial_state.is_valid():
            return None

        frontier = deque([initial_state])
        explored = set()

        while frontier:
            state = frontier.popleft()
            
            if state.is_goal():
                # Reconstruct path
                path = []
                while state.parent:
                    path.append(state.move)
                    state = state.parent
                path.reverse()
                return path

            explored.add(state)
            
            for next_state in get_next_states(state):
                if next_state not in explored and next_state not in frontier:
                    frontier.append(next_state)
        
        return None

    # Solve and speak the solution
    speak("Solving the Missionaries and Cannibals problem...")
    solution = solve()
    
    if solution:
        speak("I found a solution! Here are the steps:")
        for i, step in enumerate(solution, 1):
            speak(f"Step {i}: {step}")
            if input_type == "typing":
                print(f"Step {i}: {step}")
    else:
        speak("No solution found for the Missionaries and Cannibals problem with these parameters.")


def general_search_algorithms():
    """
    Implementation of various search algorithms with interactive problem-solving capabilities.
    Includes: BFS, DFS, UCS, A*, Greedy Best-First Search
    """
    class SearchNode:
        def __init__(self, state, parent=None, action=None, path_cost=0, heuristic=0):
            self.state = state
            self.parent = parent
            self.action = action
            self.path_cost = path_cost
            self.heuristic = heuristic
            self.depth = 0 if parent is None else parent.depth + 1

        def __lt__(self, other):
            return (self.path_cost + self.heuristic) < (other.path_cost + other.heuristic)

    class SearchProblem:
        def __init__(self, initial_state, goal_state, graph=None):
            self.initial_state = initial_state
            self.goal_state = goal_state
            self.graph = graph if graph else {}
            
        def actions(self, state):
            return self.graph.get(state, [])
            
        def result(self, state, action):
            return action
            
        def goal_test(self, state):
            return state == self.goal_state
            
        def path_cost(self, c, state1, action, state2):
            return c + 1
            
        def heuristic(self, state):
            # Manhattan distance for grid-based problems
            if isinstance(state, tuple) and isinstance(self.goal_state, tuple):
                return abs(state[0] - self.goal_state[0]) + abs(state[1] - self.goal_state[1])
            return 0

    def get_path(node):
        path = []
        while node:
            if node.action:
                path.append(node.action)
            node = node.parent
        path.reverse()
        return path

    def breadth_first_search(problem):
        node = SearchNode(problem.initial_state)
        if problem.goal_test(node.state):
            return get_path(node)
            
        frontier = deque([node])
        explored = set()
        
        while frontier:
            node = frontier.popleft()
            explored.add(node.state)
            
            for action in problem.actions(node.state):
                child = SearchNode(
                    state=problem.result(node.state, action),
                    parent=node,
                    action=action,
                    path_cost=node.path_cost + 1
                )
                
                if child.state not in explored and child not in frontier:
                    if problem.goal_test(child.state):
                        return get_path(child)
                    frontier.append(child)
        return None

    def depth_first_search(problem):
        node = SearchNode(problem.initial_state)
        if problem.goal_test(node.state):
            return get_path(node)
            
        frontier = [node]
        explored = set()
        
        while frontier:
            node = frontier.pop()
            explored.add(node.state)
            
            for action in reversed(problem.actions(node.state)):
                child = SearchNode(
                    state=problem.result(node.state, action),
                    parent=node,
                    action=action
                )
                
                if child.state not in explored and child not in frontier:
                    if problem.goal_test(child.state):
                        return get_path(child)
                    frontier.append(child)
        return None

    def uniform_cost_search(problem):
        node = SearchNode(problem.initial_state)
        frontier = []
        heapq.heappush(frontier, (node.path_cost, node))
        explored = set()
        
        while frontier:
            node = heapq.heappop(frontier)[1]
            
            if problem.goal_test(node.state):
                return get_path(node)
                
            explored.add(node.state)
            
            for action in problem.actions(node.state):
                child = SearchNode(
                    state=problem.result(node.state, action),
                    parent=node,
                    action=action,
                    path_cost=node.path_cost + problem.path_cost(node.path_cost, node.state, action, problem.result(node.state, action))
                )
                
                if child.state not in explored:
                    heapq.heappush(frontier, (child.path_cost, child))
        return None

    def a_star_search(problem):
        node = SearchNode(
            state=problem.initial_state,
            heuristic=problem.heuristic(problem.initial_state)
        )
        frontier = []
        heapq.heappush(frontier, (node.path_cost + node.heuristic, node))
        explored = set()
        
        while frontier:
            node = heapq.heappop(frontier)[1]
            
            if problem.goal_test(node.state):
                return get_path(node)
                
            explored.add(node.state)
            
            for action in problem.actions(node.state):
                child = SearchNode(
                    state=problem.result(node.state, action),
                    parent=node,
                    action=action,
                    path_cost=node.path_cost + problem.path_cost(node.path_cost, node.state, action, problem.result(node.state, action)),
                    heuristic=problem.heuristic(problem.result(node.state, action))
                )
                
                if child.state not in explored:
                    heapq.heappush(frontier, (child.path_cost + child.heuristic, child))
        return None

    def greedy_best_first_search(problem):
        node = SearchNode(
            state=problem.initial_state,
            heuristic=problem.heuristic(problem.initial_state)
        )
        frontier = []
        heapq.heappush(frontier, (node.heuristic, node))
        explored = set()
        
        while frontier:
            node = heapq.heappop(frontier)[1]
            
            if problem.goal_test(node.state):
                return get_path(node)
                
            explored.add(node.state)
            
            for action in problem.actions(node.state):
                child = SearchNode(
                    state=problem.result(node.state, action),
                    parent=node,
                    action=action,
                    heuristic=problem.heuristic(problem.result(node.state, action))
                )
                
                if child.state not in explored:
                    heapq.heappush(frontier, (child.heuristic, child))
        return None

    def get_user_input():
        speak("Please choose input method: voice or typing")
        method = takeCommand().lower()
        input_type = "voice" if "voice" in method else "typing"
        
        if input_type == "voice":
            speak("Please describe the search problem. Include start state, goal state, and connections.")
            problem_desc = takeCommand()
        else:
            problem_desc = input("Enter the search problem (format: start_state goal_state connections):\n")
            
        # Parse the input and create graph
        try:
            parts = problem_desc.split()
            initial_state = parts[0]
            goal_state = parts[1]
            graph = {}
            
            # Parse connections (format: "A-B,B-C,C-D")
            connections = parts[2].split(',')
            for conn in connections:
                src, dst = conn.split('-')
                if src not in graph:
                    graph[src] = []
                graph[src].append(dst)
                
            return SearchProblem(initial_state, goal_state, graph)
        except:
            speak("Invalid input format. Please try again.")
            return None

    def solve_with_algorithm(problem, algorithm_name):
        algorithms = {
            'bfs': (breadth_first_search, "Breadth-First Search"),
            'dfs': (depth_first_search, "Depth-First Search"),
            'ucs': (uniform_cost_search, "Uniform Cost Search"),
            'astar': (a_star_search, "A* Search"),
            'greedy': (greedy_best_first_search, "Greedy Best-First Search")
        }
        
        if algorithm_name in algorithms:
            speak(f"Solving with {algorithms[algorithm_name][1]}...")
            solution = algorithms[algorithm_name][0](problem)
            
            if solution:
                speak(f"Solution found using {algorithms[algorithm_name][1]}:")
                path_str = " -> ".join(solution)
                speak(path_str)
                print(f"Path: {path_str}")
            else:
                speak("No solution found.")
        else:
            speak("Unknown algorithm. Available algorithms: BFS, DFS, UCS, A*, Greedy")

    # Main interaction loop for search algorithms
    speak("Welcome to the general search algorithms solver.")
    while True:
        problem = get_user_input()
        if not problem:
            continue
            
        speak("Which algorithm would you like to use? Options are: BFS, DFS, UCS, A*, or Greedy")
        algorithm = takeCommand().lower()
        
        if 'bfs' in algorithm:
            solve_with_algorithm(problem, 'bfs')
        elif 'dfs' in algorithm:
            solve_with_algorithm(problem, 'dfs')
        elif 'ucs' in algorithm or 'uniform' in algorithm:
            solve_with_algorithm(problem, 'ucs')
        elif 'star' in algorithm or 'astar' in algorithm:
            solve_with_algorithm(problem, 'astar')
        elif 'greedy' in algorithm:
            solve_with_algorithm(problem, 'greedy')
        else:
            speak("Unknown algorithm specified.")
            
        speak("Would you like to try another problem or algorithm?")
        response = takeCommand().lower()
        if 'no' in response or 'exit' in response:
            break


if __name__ == "__main__":
    wishMe()
    while True:
        query = takeCommand()
        
        # Check if Jarvis is being called
        if query != "None" and "jarvis" in query.lower():
            speak("Yes, I'm here. How can I help you?")
            query = takeCommand()
        
        if query == "None":
            continue

        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "").strip()
            try:
                results = wikipedia.summary(query, sentences=2)
                speak("According to Wikipedia")
                speak(results)
            except wikipedia.exceptions.PageError:
                speak("Wikipedia page not found.")
            except wikipedia.exceptions.DisambiguationError as e:
                speak("There are multiple results for that query. Can you be more specific?")
                print(e)
            except Exception as e:
                speak("An error occurred while searching Wikipedia.")
                print(e)

        elif 'google search' in query or 'search for' in query:
            # Remove the trigger words and get the search term
            search_term = query.replace('google search', '').replace('search for', '').strip()
            if search_term:
                google_search(search_term)
            else:
                speak("What would you like me to search for?")
                search_term = takeCommand()
                if search_term != "None":
                    google_search(search_term)

        elif 'open youtube' in query:
            webbrowser.open("https://www.youtube.com")

        elif 'open whatsapp' in query:
            webbrowser.open("https://web.whatsapp.com/")

        elif 'open google' in query:
            webbrowser.open("https://www.google.com")

        elif 'open stackoverflow' in query:
            webbrowser.open("https://stackoverflow.com")

        elif 'play music' in query:
            music_dir = 'D:\\Non Critical\\songs\\Favorite Songs2'  # Replace with your music directory
            if os.path.exists(music_dir):
                songs = os.listdir(music_dir)
                if songs:
                    os.startfile(os.path.join(music_dir, songs[0]))
                else:
                    speak("No songs found in the music directory.")
            else:
                speak("Music directory not found.")
        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"The time is {strTime}")

        elif 'open code' in query:
            code_path = r"C:\Users\nishi\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Visual Studio Code\Code.exe"  # Replace with your VS Code path
            if os.path.exists(code_path):
                os.startfile(code_path)
            else:
                speak("VS Code not found at the specified path.")

        elif 'email to' in query:
            try:
                speak("Whom should I email?")
                recipient = takeCommand()
                # Basic knowledge representation: Map names to email addresses
                email_addresses = {
                    "paul": "paulsimon@gmail.com",  # Example
                    # Add more contacts here
                }
                if recipient in email_addresses:
                    to = email_addresses[recipient]
                    speak("What should I say?")
                    content = takeCommand()
                    sendEmail(to, content)
                else:
                    speak(f"Email address for {recipient} not found.")
            except Exception as e:
                print(e)
                speak("Sorry, I can't send the email right now.")

        elif 'solve tsp' in query or 'traveling salesman problem' in query:
            speak("I can help you solve the Traveling Salesman Problem.")
            distances = get_tsp_input()
            if distances:
                shortest_route, min_distance = solve_tsp(distances)
                speak("The shortest route is:")
                for i, city_index in enumerate(shortest_route):
                    speak(f"City {city_index + 1}")
                    if i < len(shortest_route) - 1:
                        speak("to")
                speak(f"with a total distance of {min_distance}")
                speak("Do you want to save this TSP data?")
                save_choice = takeCommand().lower()
                if 'yes' in save_choice:
                    save_tsp_data(distances)
            else:
                speak("Could not get valid city distances. TSP cannot be solved.")

        elif 'save tsp data' in query:
            speak("Please confirm you want to save the current TSP data.")
            save_confirm = takeCommand().lower()
            # You would need a way to access the 'distances' variable from the last TSP calculation
            # For this example, we'll assume it was stored globally or can be retrieved.
            # In a more complex application, you might need a dedicated state management.
            # For now, we'll just remind the user.
            speak("Please trigger 'save tsp data' immediately after solving a TSP problem. The current distances are not stored for saving.")

        elif 'load tsp data' in query:
            load_tsp_data()

        elif 'volume up' in query:
            control_volume("up")

        elif 'volume down' in query:
            control_volume("down")

        elif 'mute volume' in query:
            control_volume("mute")

        elif 'unmute volume' in query:
            control_volume("unmute")

        elif 'take screenshot' in query:
            take_screenshot()

        elif 'lock computer' in query:
            lock_computer()

        elif 'shutdown computer' in query:
            shutdown_computer()

        elif 'restart computer' in query:
            restart_computer()

        elif 'weather' in query:
            speak("Which city would you like to know the weather for?")
            city = takeCommand()
            if city != "None":
                get_weather(city)

        elif 'system info' in query or 'resource usage' in query:
            get_system_info()

        elif 'tell' in query and 'joke' in query:
            tell_joke()

        elif 'internet speed' in query:
            check_internet_speed()

        elif 'calculate' in query:
            speak("What would you like me to calculate?")
            calc_query = takeCommand()
            if calc_query != "None":
                calculate_expression(calc_query)

        elif 'show calendar' in query:
            get_calendar_info()

        elif 'take note' in query or 'make note' in query:
            speak("What would you like me to note down?")
            note_text = takeCommand()
            if note_text != "None":
                take_note(note_text)

        elif 'unlock computer' in query:
            unlock_computer()

        elif 'whatsapp message' in query:
            speak("Who would you like to message?")
            contact = takeCommand()
            if contact != "None":
                speak("What message should I send?")
                message = takeCommand()
                if message != "None":
                    send_whatsapp_message(contact, message)

        elif 'whatsapp call' in query:
            if 'video' in query:
                speak("Who would you like to video call?")
                contact = takeCommand()
                if contact != "None":
                    whatsapp_call(contact, video=True)
            else:
                speak("Who would you like to call?")
                contact = takeCommand()
                if contact != "None":
                    whatsapp_call(contact)

        elif 'open whatsapp chat' in query:
            speak("Which contact should I open?")
            contact = takeCommand()
            if contact != "None":
                open_whatsapp_chat(contact)

        elif 'brightness up' in query or 'increase brightness' in query:
            control_brightness("up")
            
        elif 'brightness down' in query or 'decrease brightness' in query:
            control_brightness("down")
            
        elif 'set brightness' in query:
            control_brightness("set")
            
        elif 'play spotify' in query:
            control_spotify("play")
            
        elif 'pause spotify' in query or 'stop spotify' in query:
            control_spotify("pause")
            
        elif 'next song' in query or 'next track' in query:
            control_spotify("next")
            
        elif 'previous song' in query or 'previous track' in query:
            control_spotify("previous")
            
        elif 'what\'s playing' in query or 'current song' in query:
            control_spotify("current")
            
        elif 'open' in query:
            app_name = query.replace('open', '').strip()
            if app_name:
                open_application(app_name)
                
        elif 'close' in query:
            app_name = query.replace('close', '').strip()
            if app_name:
                close_application(app_name)

        elif 'missionaries and cannibals' in query or 'solve missionaries' in query:
            solve_missionaries_cannibals()

        elif 'search algorithm' in query or 'solve search' in query or 'general search' in query:
            general_search_algorithms()

        elif 'exit' in query or 'quit' in query or 'stop' in query:
            speak("Goodbye!")
            break