import tkinter as tk
from tkinter import messagebox, simpledialog
import pandas as pd
import random

class QuizApp:
    def __init__(self, master):
        """
        Inicializa la aplicación del quiz de Harry Potter.
        
        :param master: La ventana principal de Tkinter.
        """
        self.master = master
        self.master.title("Harry Potter Quiz")
        self.score = 0
        self.questions_asked = set()  # Almacena preguntas ya hechas como tuplas
        self.current_question = None
        self.player_name = None
        self.timer_running = False  # Controla si el temporizador está en ejecución
        self.option_buttons = []  # Inicializamos option_buttons aquí para asegurar que siempre exista
        self.time_label = None  # Inicializamos time_label para controlar un solo temporizador
        self.question_counter = 0  # Contador de preguntas

        # Cargar preguntas desde Excel
        self.questions = self.load_questions_from_excel('questions.xlsx')
        
        self.label = tk.Label(master, text="Bienvenido al Quiz de Harry Potter!")
        self.label.pack(pady=10)

        self.start_button = tk.Button(master, text="Jugar", command=self.get_player_name)
        self.start_button.pack(pady=5)

        self.ranking_button = tk.Button(master, text="Ranking", command=self.show_ranking)
        self.ranking_button.pack(pady=5)

    def load_questions_from_excel(self, filename):
        """
        Carga las preguntas desde un archivo Excel.

        :param filename: El nombre del archivo Excel con las preguntas.
        :return: Lista de diccionarios con información de cada pregunta.
        """
        try:
            df = pd.read_excel(filename)
            questions = []
            for _, row in df.iterrows():
                questions.append({
                    "question": row['Question'],
                    "options": [row['Option1'], row['Option2'], row['Option3']],
                    "correct": row['CorrectAnswer']
                })
            return questions
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar preguntas: {e}")
            return []

    def get_player_name(self):
        """
        Solicita el nombre del jugador para iniciar el juego.
        """
        self.player_name = simpledialog.askstring("Nombre del Jugador", "Por favor, introduce tu seudónimo:")
        if self.player_name:
            self.start_quiz()
        else:
            messagebox.showwarning("Nombre Requerido", "Debes proporcionar un seudónimo para jugar.")

    def start_quiz(self):
        """
        Inicia el quiz, ocultando los botones de inicio y ranking.
        """
        self.label.config(text="")
        self.start_button.pack_forget()
        self.ranking_button.pack_forget()
        self.score = 0
        self.questions_asked.clear()
        self.question_counter = 0
        self.ask_question()

    def ask_question(self):
        """
        Presenta una nueva pregunta al jugador, gestionando el temporizador y las respuestas.
        Este método asegura que los botones de opción siempre están presentes y actualizados para cada pregunta.
        """
        self.question_counter += 1
        if self.question_counter > 10:
            self.end_game()
            return

        while True:
            self.current_question = random.choice(self.questions)
            question_tuple = (self.current_question["question"], self.current_question["correct"])
            if question_tuple not in self.questions_asked:
                break
        self.questions_asked.add(question_tuple)

        self.label.config(text=f"Pregunta {self.question_counter}/10: {self.current_question['question']}")
        
        # Asegúrate de que los botones de opción existen para cada pregunta
        for button in self.option_buttons:
            button.pack_forget()
        self.option_buttons.clear()

        for i, option in enumerate(self.current_question["options"]):
            button = tk.Button(self.master, text=option, bg='white', command=lambda opt=option: self.check_answer(opt))
            button.pack(pady=5)
            self.option_buttons.append(button)

        # Inicia el temporizador sin usar threading, asegurando que solo haya un contador
        if self.time_label:
            self.time_label.pack_forget()
        self.timer = 30
        self.time_label = tk.Label(self.master, text=f"Tiempo restante: {self.timer}")
        self.time_label.pack()
        self.start_countdown()

    def start_countdown(self):
        """
        Inicia el temporizador para la pregunta actual usando `after` para actualizaciones.
        """
        self.timer_running = True
        self.countdown()

    def countdown(self):
        """
        Maneja la cuenta regresiva del temporizador de la pregunta, usando únicamente `after` para ser thread-safe.
        Este método asegura que el temporizador mantenga una velocidad constante.
        """
        if self.timer_running:
            if self.timer > 0:
                self.timer -= 1
                self.time_label.config(text=f"Tiempo restante: {self.timer}")
                self.master.after(1000, self.countdown)  # Mantiene la velocidad constante con 1000 ms por iteración
            else:
                self.timer_running = False
                self.check_answer(None)  # Tiempo agotado

    def check_answer(self, answer):
        """
        Verifica la respuesta del jugador y gestiona el puntaje.

        :param answer: La respuesta seleccionada por el jugador o None si el tiempo expiró.
        """
        try:
            # Asegura que todos los botones sean olvidados para evitar duplicados
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
            self.timer_running = False  # Aseguramos que el temporizador esté detenido
            if self.question_counter < 10:
                self.ask_question()
            else:
                self.end_game()
        except Exception as e:
            messagebox.showerror("Error", f"Un error ocurrió durante la verificación de la respuesta: {e}")

    def end_game(self):
        """
        Termina el juego, muestra el puntaje, guarda en el ranking y vuelve a la pantalla inicial.
        """
        messagebox.showinfo("Fin del Juego", f"Juego terminado. Tu puntuación final es: {self.score}")
        self.save_score_to_ranking()  # Guarda solo cuando se completen todas las preguntas
        
        # Vuelve a la pantalla inicial
        self.label.config(text="Bienvenido al Quiz de Harry Potter!")
        self.start_button.pack(pady=5)
        self.ranking_button.pack(pady=5)
        self.player_name = None  # Reset para el siguiente jugador

    def save_score_to_ranking(self):
        """
        Guarda la puntuación del jugador en el archivo de ranking.
        """
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
        """
        Muestra el ranking actual de jugadores.
        """
        try:
            df = pd.read_excel('ranking.xlsx')
            if df.empty:
                messagebox.showinfo("Ranking", "No hay puntuaciones en el ranking todavía.")
            else:
                ranking_text = "\n".join([f"{row['Name']}: {row['Score']}" for _, row in df.iterrows()])
                messagebox.showinfo("Ranking", ranking_text)
        except FileNotFoundError:
            messagebox.showinfo("Ranking", "No hay puntuaciones en el ranking todavía.")

# Inicialización de la aplicación
root = tk.Tk()
app = QuizApp(root)
root.mainloop()