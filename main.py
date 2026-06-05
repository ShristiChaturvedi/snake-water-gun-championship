import tkinter as tk
from tkinter import messagebox
import random
import win32com.client

# Voice Engine
speaker = win32com.client.Dispatch("SAPI.SpVoice")

def speak(text):
    speaker.Speak(text)

# Scores
user_score = 0
computer_score = 0
draws = 0

# Main Window
root = tk.Tk()
root.title("Snake Water Gun Championship")
root.geometry("700x500")
root.resizable(False, False)

# Title
title = tk.Label(
    root,
    text="🐍 Snake Water Gun Championship 🔫",
    font=("Arial", 22, "bold")
)
title.pack(pady=20)

# Scoreboard
score_label = tk.Label(
    root,
    text="You: 0   Computer: 0   Draws: 0",
    font=("Arial", 16)
)
score_label.pack()

# Result Label
result_label = tk.Label(
    root,
    text="Choose your move",
    font=("Arial", 16)
)
result_label.pack(pady=20)

# Computer Choice
computer_choice_label = tk.Label(
    root,
    text="🤖 Computer Choice: ?",
    font=("Arial", 18)
)
computer_choice_label.pack(pady=10)

def update_score():
    score_label.config(
        text=f"You: {user_score}   Computer: {computer_score}   Draws: {draws}"
    )

def reset_game():
    global user_score, computer_score, draws

    user_score = 0
    computer_score = 0
    draws = 0

    update_score()

    result_label.config(text="Choose your move")
    computer_choice_label.config(text="🤖 Computer Choice: ?")

def play(user_choice):
    global user_score, computer_score, draws

    computer_choice = random.choice(["Snake", "Water", "Gun"])

    emoji = {
        "Snake": "🐍",
        "Water": "💧",
        "Gun": "🔫"
    }

    computer_choice_label.config(
        text=f"🤖 Computer Choice: {emoji[computer_choice]} {computer_choice}"
    )

    if user_choice == computer_choice:

        draws += 1
        result_text = "🤝 Draw!"
        speak("It's a draw")

    elif (
        (user_choice == "Snake" and computer_choice == "Water") or
        (user_choice == "Water" and computer_choice == "Gun") or
        (user_choice == "Gun" and computer_choice == "Snake")
    ):

        user_score += 1
        result_text = "🎉 You Win!"
        speak("Congratulations. You Win")

    else:

        computer_score += 1
        result_text = "😢 Computer Wins!"
        speak("Computer Wins")

    result_label.config(
        text=f"You: {user_choice} | Computer: {computer_choice}\n{result_text}"
    )

    update_score()

    # Champion at 5 points
    if user_score == 5:

        speak(
            f"Amazing! You are the champion of Snake Water Gun Championship. "
            f"You scored {user_score} points while the computer scored {computer_score} points."
        )

        messagebox.showinfo(
            "Champion",
            "🏆 YOU ARE THE CHAMPION! 🏆"
        )

        reset_game()

    elif computer_score == 5:

        speak(
            f"The computer is the champion with {computer_score} points. "
            f"You scored {user_score} points."
        )

        messagebox.showinfo(
            "Champion",
            "💀 COMPUTER IS THE CHAMPION! 💀"
        )

        reset_game()

# Buttons Frame
button_frame = tk.Frame(root)
button_frame.pack(pady=30)

snake_btn = tk.Button(
    button_frame,
    text="🐍 Snake",
    font=("Arial", 16),
    width=12,
    command=lambda: play("Snake")
)
snake_btn.grid(row=0, column=0, padx=10)

water_btn = tk.Button(
    button_frame,
    text="💧 Water",
    font=("Arial", 16),
    width=12,
    command=lambda: play("Water")
)
water_btn.grid(row=0, column=1, padx=10)

gun_btn = tk.Button(
    button_frame,
    text="🔫 Gun",
    font=("Arial", 16),
    width=12,
    command=lambda: play("Gun")
)
gun_btn.grid(row=0, column=2, padx=10)

# Exit Button
exit_btn = tk.Button(
    root,
    text="❌ Exit Game",
    font=("Arial", 14),
    command=root.destroy
)
exit_btn.pack(pady=20)

# Welcome Voice
speak("Welcome to Snake Water Gun Championship")

root.mainloop()
