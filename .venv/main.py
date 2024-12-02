import win32com.client
import pyttsx3
import speech_recognition as sr
import webbrowser
from langchain_ollama import OllamaLLM
import mtranslate
from gtts import gTTS
import pygame
import tempfile
import os

# AI Initializing
model = OllamaLLM(model="llama3.2:1b")



def say(text):
    try:
        # Initialize pygame mixer (for audio playback)
        pygame.mixer.init()

        # Translate the text to Urdu
        text = mtranslate.translate(text, to_language="ur", from_language="en")

        # Create a temporary file to save the audio, ensuring it stays open
        with tempfile.NamedTemporaryFile(delete=False, mode='wb', suffix=".mp3") as temp_file:
            temp_filename = temp_file.name  # Get the name/path of the temp file

            # Convert text to speech (in Urdu)
            tts = gTTS(text, lang='ur')
            tts.save(temp_filename)  # Save the speech to the temporary file

        # Load and play the saved audio
        pygame.mixer.music.load(temp_filename)
        pygame.mixer.music.set_volume(1.0)  # Ensure full volume
        pygame.mixer.music.play()

        # Wait until the audio finishes playing (block until done)
        while pygame.mixer.music.get_busy():  # This keeps the program from continuing
            pygame.time.Clock().tick(10)  # Wait and check every 10ms

        # Optionally, clean up the temporary file after playback
        os.remove(temp_filename)

    except pygame.error as e:
        print(f"pygame error: {e}")
        return "Audio playback error occurred!"
    except Exception as e:
        print(f"Error: {str(e)}")
        return "Sorry! Some error occurred!"





def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("recognizing...")
        r.pause_threshold = 1
        r.adjust_for_ambient_noise(source)


        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-pk")
            query = mtranslate.translate(query, to_language="en-pk")
            print(f"User said: {query}")
            return query

        except Exception as e:
            print(f"Error: {str(e)}")
            return "Sorry! some error occurred!"


def get_ollama_response(query):

    response = model.invoke(input=f"{query} . answer in those tone in which question has been asked")


    response = response.replace("Meta", "Naveed").replace("Llama", "Naveed AI")

    return response


if __name__ == '__main__':
    print("Initializing Agha AI Bot...")
    say("I am Agha AI Bot")

    while True:
        print("Listening...")

        query = takeCommand()

        if query.lower() in ['sorry! some error occurred!', '']:
            continue  # Skip if the query was an error

        # Check if the query is asking to open a website
        sites = [['youtube', "https://www.youtube.com"],
                 ['wikipedia', "https://www.wikipedia.com"],
                 ['google', "https://www.google.com"],
                 ['naveed github', "https://www.github.com/agha-naveed"]]

        for site in sites:
            if f"open {site[0]}".lower() in query.lower():
                print(f"Opening {site[0]}...")
                webbrowser.open(site[1])

                say(f"Opening {site[0]}")
                break

        response = get_ollama_response(query)
        print(f"AI Model says: {response}")

        say(response)
