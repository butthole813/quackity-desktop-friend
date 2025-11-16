import tkinter as tk
from PIL import Image, ImageTk
import win32com.client
import os
import random
import sys
from tkinter import messagebox
from datetime import date

# ---------------------------
# PNG path fix for EXE
# ---------------------------
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

image_path = os.path.join(base_path, "quackster.png")
data_file = os.path.join(base_path, "quackster_last.txt")

# ---------------------------
# Microsoft David TTS
# ---------------------------
speaker = win32com.client.Dispatch("SAPI.SpVoice")

# ---------------------------
# Tkinter window
# ---------------------------
root = tk.Tk()
root.title("Quackster Debug")
root.attributes('-topmost', True)

img = Image.open(image_path)
tk_img = ImageTk.PhotoImage(img)

canvas = tk.Canvas(root, width=img.width, height=img.height, highlightthickness=0)
canvas.pack()
canvas.create_image(0, 0, anchor='nw', image=tk_img)

# ---------------------------
# Eye definitions
# ---------------------------
eye_data = [
    {"rect": [15, 15, 35, 35]},  # Left eye
    {"rect": [50, 15, 70, 35]},  # Right eye
]

# Pupils: black squares
for eye in eye_data:
    x1, y1, x2, y2 = eye["rect"]
    size = min(x2 - x1, y2 - y1) // 2  # square size
    cx = (x1 + x2) // 2
    cy = (y1 + y2) // 2
    eye["pupil"] = canvas.create_rectangle(cx - size//2, cy - size//2, cx + size//2, cy + size//2, fill='black', outline='black')

# ---------------------------
# Eye tracking
# ---------------------------
def update_eyes():
    mouse_x, mouse_y = root.winfo_pointerxy()
    root_x = root.winfo_rootx()
    root_y = root.winfo_rooty()

    for eye in eye_data:
        x1, y1, x2, y2 = eye["rect"]
        cx = (x1 + x2) / 2
        cy = (y1 + y2) / 2
        width = (x2 - x1) / 2
        height = (y2 - y1) / 2

        dx = mouse_x - (root_x + cx)
        dy = mouse_y - (root_y + cy)

        if dx > width: dx = width
        if dx < -width: dx = -width
        if dy > height: dy = height
        if dy < -height: dy = -height

        size = (x2 - x1) // 2
        canvas.coords(eye["pupil"], cx + dx - size//2, cy + dy - size//2, cx + dx + size//2, cy + dy + size//2)

    root.after(30, update_eyes)

# ---------------------------
# Blinking
# ---------------------------
def blink():
    for eye in eye_data:
        coords = canvas.coords(eye["pupil"])
        cx = (coords[0] + coords[2]) / 2
        cy = (coords[1] + coords[3]) / 2
        width = coords[2] - coords[0]
        height = coords[3] - coords[1]
        new_height = height * 1.5  # rectangle for blink
        canvas.coords(eye["pupil"], cx - width/2, cy - new_height/2, cx + width/2, cy + new_height/2)

    root.after(150, restore_pupils)
    root.after(random.randint(2000, 5000), blink)

def restore_pupils():
    for eye in eye_data:
        x1, y1, x2, y2 = eye["rect"]
        size = min(x2 - x1, y2 - y1) // 2
        cx = (x1 + x2) / 2
        cy = (y1 + y2) / 2
        canvas.coords(eye["pupil"], cx - size//2, cy - size//2, cx + size//2, cy + size//2)

# ---------------------------
# Drag functionality (middle-click)
# ---------------------------
def start_drag(event):
    global drag_data
    drag_data = {'x': event.x, 'y': event.y}

def do_drag(event):
    x = root.winfo_x() + event.x - drag_data['x']
    y = root.winfo_y() + event.y - drag_data['y']
    root.geometry(f"+{x}+{y}")

canvas.bind("<Button-2>", start_drag)
canvas.bind("<B2-Motion>", do_drag)

# ---------------------------
# Left-click to speak
# ---------------------------
def speak(event):
    phrases = ["Hello!", "Hi there!", "Greetings!", "Quack!"]
    speaker.Speak(random.choice(phrases))

canvas.bind("<Button-1>", speak)

# ---------------------------
# Right-click to remove
# ---------------------------
def ask_remove(event):
    answer = messagebox.askquestion("Remove Quackster?", "Do you want to remove Quackster from the screen? (Y/N)")
    if answer == "yes":
        speaker.Speak("Bye.")
        speaker.Speak("See you later!")
        root.destroy()
    else:
        stay_phrases = ["Okay, I’ll stay with you!", "Yay! I’m staying!", "I won’t leave, promise!", "Quack! I’m not going anywhere!"]
        speaker.Speak(random.choice(stay_phrases))

canvas.bind("<Button-3>", ask_remove)

# ---------------------------
# Startup greeting based on last run
# ---------------------------
today = str(date.today())
if os.path.exists(data_file):
    with open(data_file, "r") as f:
        last_run = f.read().strip()
    if last_run == today:
        speaker.Speak("It's you again!")
        speaker.Speak("Hi again!")
else:
    last_run = ""

with open(data_file, "w") as f:
    f.write(today)

# ---------------------------
# Start loops
# ---------------------------
update_eyes()
blink()
root.mainloop()
