import os
import queue
import json
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
import pyaudio
from vosk import Model, KaldiRecognizer

class SpeechToTextApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Real-Time Speech-to-Text")
        self.root.geometry("800x630")
        self.root.resizable(False, False)

        self.recording = False
        self.q = queue.Queue()
        self.model = Model(r"C:\Users\infinite\Downloads\vosk-model-small-hi-0.22")  # Replace with the path to your Vosk model
        self.create_widgets()

    def create_widgets(self):
        menu = tk.Menu(self.root)
        self.root.config(menu=menu)

        file_menu = tk.Menu(menu, tearoff=0)
        file_menu.add_command(label="Save", command=self.save_text)
        menu.add_cascade(label="File", menu=file_menu)

        toolbar = tk.Frame(self.root, bg="#f0f0f0")
        toolbar.pack(side=tk.TOP, fill=tk.X)

        self.recording_button = tk.Button(toolbar, text="Start Recording", command=self.toggle_recording)
        self.recording_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.status_label = tk.Label(toolbar, text="Ready", fg="green")
        self.status_label.pack(side=tk.RIGHT, padx=5, pady=5)

        self.text_box = tk.Text(self.root, font=("Arial", 14), wrap="word")
        self.text_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

    def toggle_recording(self):
        if not self.recording:
            self.start_recording()
        else:
            self.stop_recording()

    def start_recording(self):
        self.recording = True
        self.status_label.config(text="Listening...", fg="red")
        self.record_thread = threading.Thread(target=self.record_audio)
        self.record_thread.start()

    def stop_recording(self):
        self.recording = False
        self.status_label.config(text="Ready", fg="green")
        if self.record_thread.is_alive():
            self.record_thread.join()

    def record_audio(self):
        audio_format = pyaudio.paInt16
        channels = 1
        rate = 16000
        chunk = 1024

        audio_interface = pyaudio.PyAudio()
        audio_stream = audio_interface.open(
            format=audio_format,
            channels=channels,
            rate=rate,
            input=True,
            frames_per_buffer=chunk,
            stream_callback=self.callback
        )

        while self.recording:
            self.root.update()

        audio_stream.stop_stream()
        audio_stream.close()
        audio_interface.terminate()

    def callback(self, in_data, frame_count, time_info, status):
        if not self.recording:
            return (None, pyaudio.paComplete)
        
        self.q.put(in_data)
        recognizer = KaldiRecognizer(self.model, 16000)
        
        if recognizer.AcceptWaveform(in_data):
            result = json.loads(recognizer.Result())
            self.update_text(result.get('text', ''))
        else:
            partial_result = json.loads(recognizer.PartialResult())
            self.update_text(partial_result.get('partial', ''), final=False)
        
        return (None, pyaudio.paContinue)

    def update_text(self, text, final=True):
        if final:
            self.text_box.insert(tk.END, text + ' ')
        else:
            current_text = self.text_box.get("1.0", tk.END).strip()
            words = current_text.split()
            if words:
                words[-1] = text
                new_text = ' '.join(words)
            else:
                new_text = text
            self.text_box.delete("1.0", tk.END)
            self.text_box.insert(tk.END, new_text)

    def save_text(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
        if file_path:
            self.save_text_to_docx(file_path)

    def save_text_to_docx(self, file_path):
        doc = Document()
        doc.add_paragraph(self.text_box.get("1.0", tk.END))
        try:
            doc.save(file_path)
            messagebox.showinfo("Success", "Text saved as Word document.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save document: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = SpeechToTextApp(root)
    root.mainloop()
