import threading

class SpeechService:
    def __init__(self):
        self.available = False
        self.error = None
        try:
            import pyttsx3
            self.pyttsx3 = pyttsx3
            self.available = True
        except Exception as e:
            self.pyttsx3 = None
            self.error = e

    def speak(self, text: str):
        if not text:
            return
        if not self.available:
            raise RuntimeError("pyttsx3 is not installed. Run: pip install pyttsx3")

        def run():
            engine = self.pyttsx3.init()
            engine.say(text)
            engine.runAndWait()
            engine.stop()

        threading.Thread(target=run, daemon=True).start()
