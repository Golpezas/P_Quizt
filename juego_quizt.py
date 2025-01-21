import tkinter as tk
from tkinter import messagebox, simpledialog
import pandas as pd
import random
import time
from openpyxl import Workbook

class QuizApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Harry Potter Quiz")
        self.score = 0
        self.questions_asked = set()  # Ahora es un set de tuplas
        self.current_question = None
        self.player_name = None

        # Cargar preguntas desde Excel
        self.questions = self.load_questions_from_excel('questions.xlsx')
        
        self.label = tk.Label(master, text="Bienvenido al Quiz de Harry Potter!")
        self.label.pack(pady=10)

        self.start_button = tk.Button(master, text="Jugar", command=self.get_player_name)
        self.start_button.pack(pady=5)

        self.ranking_button = tk.Button(master, text="Ranking", command=self.show_ranking)
        self.ranking_button.pack(pady=5)

    def load_questions_from_excel(self, filename):
        df = pd.read_excel(filename)
        questions = []
        for _, row in df.iterrows():
            questions.append({
                "question": row['Question'],
                "options": [row['Option1'], row['Option2'], row['Option3']],
                "correct": row['CorrectAnswer']
            })
        return questions

    def get_player_name(self):
        self.player_name = simpledialog.askstring("Nombre del Jugador", "Por favor, introduce tu seudónimo:")
        if self.player_name:
            self.start_quiz()
        else:
            messagebox.showwarning("Nombre Requerido", "Debes proporcionar un seudónimo para jugar.")

    def start_quiz(self):
        self.label.config(text="")
        self.start_button.pack_forget()
        self.ranking_button.pack_forget()
        self.ask_question()

    def ask_question(self):
        if len(self.questions_asked) >= 10:
            self.end_game()
            return

        while True:
            self.current_question = random.choice(self.questions)
            # Crear una tupla inmutable para usar como clave
            question_tuple = (self.current_question["question"], self.current_question["correct"])
            if question_tuple not in self.questions_asked:
                break
        self.questions_asked.add(question_tuple)

        self.label.config(text=self.current_question["question"])
        
        self.options = []
        for i, option in enumerate(self.current_question["options"]):
            button = tk.Button(self.master, text=option, command=lambda opt=option: self.check_answer(opt))
            button.pack(pady=5)
            self.options.append(button)

        self.timer = 30
        self.time_label = tk.Label(self.master, text=f"Tiempo restante: {self.timer}")
        self.time_label.pack()
        self.countdown()

    def countdown(self):
        if self.timer > 0:
            self.timer -= 1
            self.time_label.config(text=f"Tiempo restante: {self.timer}")
            self.master.after(1000, self.countdown)
        else:
            self.check_answer(None)  # Tiempo agotado

    def check_answer(self, answer):
        for button in self.options:
            button.pack_forget()
        self.time_label.pack_forget()

        if answer == self.current_question["correct"]:
            points = self.timer if self.timer > 0 else 0
            self.score += points
            messagebox.showinfo("Correcto", f"¡Correcto! Ganaste {points} puntos.")
        else:
            messagebox.showinfo("Incorrecto", f"Incorrecto. La respuesta correcta era: {self.current_question['correct']}")

        if len(self.questions_asked) < 10:
            self.ask_question()
        else:
            self.end_game()

    def end_game(self):
        messagebox.showinfo("Fin del Juego", f"Juego terminado. Tu puntuación final es: {self.score}")
        self.save_score_to_ranking()
        self.master.quit()

    def save_score_to_ranking(self):
        try:
            df = pd.read_excel('ranking.xlsx')
        except FileNotFoundError:
            df = pd.DataFrame(columns=['Name', 'Score'])
        
        new_row = pd.DataFrame({'Name': [self.player_name], 'Score': [self.score]})
        df = pd.concat([df, new_row], ignore_index=True)
        df = df.sort_values(by='Score', ascending=False)
        df.to_excel('ranking.xlsx', index=False)

    def show_ranking(self):
        try:
            df = pd.read_excel('ranking.xlsx')
            if df.empty:
                messagebox.showinfo("Ranking", "No hay puntuaciones en el ranking todavía.")
            else:
                ranking_text = "\n".join([f"{row['Name']}: {row['Score']}" for _, row in df.iterrows()])
                messagebox.showinfo("Ranking", ranking_text)
        except FileNotFoundError:
            messagebox.showinfo("Ranking", "No hay puntuaciones en el ranking todavía.")

root = tk.Tk()
app = QuizApp(root)
root.mainloop()