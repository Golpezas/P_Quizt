import tkinter as tk
from tkinter import messagebox, simpledialog
import pandas as pd
import random
import time
import threading

class QuizApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Harry Potter Quiz")
        self.score = 0
        self.questions_asked = set()  # Ahora es un set de tuplas
        self.current_question = None
        self.player_name = None
        self.timer_running = False  # Nuevo atributo para controlar el estado del temporizador

        # Cargar preguntas desde Excel
        self.questions = self.load_questions_from_excel('questions.xlsx')
        
        self.label = tk.Label(master, text="Bienvenido al Quiz de Harry Potter!")
        self.label.pack(pady=10)

        self.start_button = tk.Button(master, text="Jugar", command=self.get_player_name)
        self.start_button.pack(pady=5)

        self.ranking_button = tk.Button(master, text="Ranking", command=self.show_ranking)
        self.ranking_button.pack(pady=5)

    # ... (métodos load_questions_from_excel, get_player_name, start_quiz iguales)

    def ask_question(self):
        if len(self.questions_asked) >= 10:
            self.end_game()
            return

        while True:
            self.current_question = random.choice(self.questions)
            question_tuple = (self.current_question["question"], self.current_question["correct"])
            if question_tuple not in self.questions_asked:
                break
        self.questions_asked.add(question_tuple)

        self.label.config(text=self.current_question["question"])
        
        if not hasattr(self, 'option_buttons'):
            self.option_buttons = []
            for i in range(3):
                button = tk.Button(self.master, command=lambda i=i: self.check_answer(self.current_question["options"][i]))
                self.option_buttons.append(button)
                button.pack(pady=5)

        for button, option in zip(self.option_buttons, self.current_question["options"]):
            button.config(text=option, bg='white')

        if hasattr(self, 'countdown_thread') and self.countdown_thread.is_alive():
            self.timer_running = False  # Esto detiene el temporizador
            self.countdown_thread.join()  # Espera a que termine el hilo anterior

        self.timer = 30
        self.time_label = tk.Label(self.master, text=f"Tiempo restante: {self.timer}")
        self.time_label.pack()
        self.start_countdown()

    def start_countdown(self):
        self.timer_running = True
        self.countdown_thread = threading.Thread(target=self.countdown)
        self.countdown_thread.start()

    def countdown(self):
        for remaining in range(self.timer, 0, -1):
            if not self.timer_running:
                return
            time.sleep(1)
            self.master.after(0, lambda r=remaining: self.time_label.config(text=f"Tiempo restante: {r}"))
        if self.timer_running:
            self.master.after(0, self.check_answer, None)  # Tiempo agotado
        self.timer_running = False

    def check_answer(self, answer):
        if hasattr(self, 'countdown_thread') and self.countdown_thread.is_alive():
            self.timer_running = False  # Esto detiene el temporizador
            self.countdown_thread.join()  # Asegura que el hilo termine antes de proceder

        for button in self.option_buttons:
            button.pack_forget()
        self.time_label.pack_forget()

        if answer == self.current_question["correct"]:
            points = self.timer if self.timer > 0 else 0
            self.score += points
            messagebox.showinfo("Correcto", f"¡Correcto! Ganaste {points} puntos.")
            self.label.config(fg='green')  # Visual feedback
        else:
            messagebox.showinfo("Incorrecto", f"Incorrecto. La respuesta correcta era: {self.current_question['correct']}")
            self.label.config(fg='red')    # Visual feedback

        self.label.config(fg='black')  # Reset color
        if len(self.questions_asked) < 10:
            self.ask_question()
        else:
            self.end_game()

    def end_game(self):
        messagebox.showinfo("Fin del Juego", f"Juego terminado. Tu puntuación final es: {self.score}")
        self.save_score_to_ranking()  # Mover aquí para guardar solo cuando se completen todas las preguntas
        self.master.quit()

    def save_score_to_ranking(self):
        try:
            df = pd.read_excel('ranking.xlsx')
        except FileNotFoundError:
            df = pd.DataFrame(columns=['Name', 'Score'])
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo de ranking: {e}")
            return

        new_row = pd.DataFrame({'Name': [self.player_name], 'Score': [self.score]})
        df = pd.concat([df, new_row], ignore_index=True)
        df = df.sort_values(by='Score', ascending=False)

        try:
            df.to_excel('ranking.xlsx', index=False)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el ranking: {e}")

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