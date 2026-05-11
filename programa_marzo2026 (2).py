import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
import pandas as pd
import random
import os  
import sys 
import platform 
import urllib.request
import urllib.parse
import json
import threading
import datetime

# Definición de la Paleta de Colores
BG_PRINCIPAL = '#f0f8ff'  # Blanco Azulado (Azul muy suave)
BG_ACENTO = '#dbe9f6'    # Azul Claro para Frames
FG_TITULO = '#1e90ff'    # Azul Intenso (Moderno)
BG_BOTON = '#ff6347'     # Rojo Tomate (Acento de acción)
FG_TEXTO = '#333333'     # Gris Oscuro

GOOGLE_SHEETS_URL = "https://script.google.com/macros/s/AKfycbzl4dWOvaZ9UPzSWxzM6mv5oh0vHyTB5qy8fDtVVpaYQfXLpK6nHxFt4m_jQ9iQlE1Z/exec"

def enviar_a_sheets(datos):
    """Envía los datos al Google Sheet en un hilo separado para no bloquear la interfaz."""
    def _enviar():
        try:
            payload = json.dumps(datos).encode('utf-8')
            req = urllib.request.Request(
                GOOGLE_SHEETS_URL,
                data=payload,
                headers={'Content-Type': 'application/json'}
            )
            urllib.request.urlopen(req, timeout=10)
        except Exception:
            pass  # Si falla el envío, no interrumpe el programa
    threading.Thread(target=_enviar, daemon=True).start()

class EvaluacionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Evaluación de Psicología")
        self.root.configure(bg=BG_PRINCIPAL)
        
        # Inicia la ventana maximizada
        self.root.state('zoomed')
        
        # Protocolo de cierre con confirmación
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.data = {}
        self.answers = []
        self.logo_image = None
        self.current_image = None
        self.create_welcome_screen()

    # --- Funciones de Archivo y Cierre ---

    def open_excel_file(self, filename='resultado_evaluacion_conductual.xlsx'):
        """Abre el archivo Excel guardado usando el comando del sistema operativo."""
        if not os.path.exists(filename):
            messagebox.showinfo("Error de Archivo", f"El archivo de resultados '{filename}' no existe.")
            return

        try:
            current_platform = platform.system()
            if current_platform == "Windows":
                os.startfile(filename)
            elif current_platform == "Darwin":
                os.system(f"open '{filename}'")
            elif current_platform == "Linux":
                os.system(f"xdg-open '{filename}'")
            else:
                messagebox.showinfo("Información de Guardado", f"El archivo se guardó como '{filename}', pero no se pudo abrir automáticamente en tu sistema operativo.")

        except Exception as e:
            # Este error es al intentar abrir el archivo, no al guardar
            messagebox.showinfo("Error al Abrir", f"El archivo se guardó, pero hubo un error al intentar abrirlo: {e}. Por favor, ábrelo manualmente.")

    def save_and_open_excel(self):
        """
        Guarda todos los datos en Excel en formato simplificado (Pregunta/Respuesta) 
        y abre el archivo.
        """
        
        result_data = self.data.copy()
        
        # Formato simplificado: Pregunta/Respuesta
        for idx, ans in enumerate(self.answers, start=1):
            result_data[f"Pregunta {idx}"] = ans['Pregunta']
            result_data[f"Respuesta {idx}"] = ans['Respuesta Seleccionada']

        df = pd.DataFrame([result_data])
        filename = 'resultado_evaluacion_conductual.xlsx'
        
        try:
            df.to_excel(filename, index=False)
            self.open_excel_file(filename)
        except Exception as e:
            # Manejo de errores mejorado
            error_message = f"No se pudo guardar el archivo Excel. Asegúrate de que el archivo '{filename}' NO esté abierto por otra aplicación."
            if "No such file or directory" in str(e) or "No module named 'openpyxl'" in str(e):
                 error_message += "\n\nTambién verifica que tengas la librería 'openpyxl' instalada: pip install openpyxl"
            else:
                error_message += f"\n\nDetalle del error: {e}"

            messagebox.showerror("Error de Guardado", error_message)

    def on_closing(self):
        """Maneja el evento de cierre de la ventana con confirmación y guarda el progreso."""
        if messagebox.askyesno("Confirmar Cierre", "¿Estás seguro de que deseas cerrar el programa? Se guardarán los datos recopilados hasta ahora."):
            # Solo guardamos si hay datos de demografía, autoevaluación o respuestas.
            if self.data or self.answers:
                self.save_and_open_excel()
            
            self.root.destroy()

    # --- Métodos de Interfaz (Omitidos para brevedad, no hay cambios en la lógica) ---
    # ... (Se mantiene el código de create_welcome_screen, create_consentimiento_window, etc.)
    # Se añade todo el código de las vistas aquí.
    
    def clear_window(self):
        for widget in self.root.winfo_children():
            widget.destroy()
        self.current_image = None 

    def create_welcome_screen(self):
        self.clear_window()
        
        # Frame principal para centrar contenido
        main_frame = tk.Frame(self.root, bg=BG_PRINCIPAL)
        main_frame.pack(pady=50, padx=50)

        try:
            image = Image.open("image.png")
            image = image.resize((200, 200)) 
            self.logo_image = ImageTk.PhotoImage(image)
            tk.Label(main_frame, image=self.logo_image, bg=BG_PRINCIPAL).pack(pady=20)
        except Exception:
            pass

        tk.Label(main_frame, text="¡Bienvenido/a! 👋", font=("Arial", 24, 'bold'), bg=BG_PRINCIPAL, fg=FG_TITULO).pack(pady=15)
        instrucciones = "Muchas gracias por tu participación. Nos gustaría que nos ayudaras a contestar algunas preguntas sobre tu formación psicológica. Te pedimos que contestes con honestidad, sin consultar fuentes externas. Asimismo, si tienes alguna duda, pregunta al investigador porque una vez iniciada la evaluación no podrás hacerlo."
        tk.Label(main_frame, text=instrucciones, wraplength=700, justify='center', bg=BG_PRINCIPAL, fg=FG_TEXTO, font=('Arial', 14)).pack(padx=30, pady=30)
        
        tk.Button(main_frame, text="Entendido ✅", bg=BG_BOTON, fg='white', activebackground='#ff8c73', 
                  font=('Arial', 14, 'bold'), command=self.create_consentimiento_window, 
                  width=25, height=2, bd=0).pack(pady=40)

    def create_consentimiento_window(self):
        self.clear_window()
        main_frame = tk.Frame(self.root, bg=BG_PRINCIPAL)
        main_frame.pack(pady=50, padx=50)

        tk.Label(main_frame, text="Consentimiento Informado", font=("Arial", 20, 'bold'), bg=BG_PRINCIPAL, fg=FG_TITULO).pack(pady=20)
        texto = "He sido informado del objetivo de la investigación, de mi participación, y del manejo confidencial de mis datos. Estoy de acuerdo en participar."
        tk.Label(main_frame, text=texto, wraplength=700, justify='center', bg=BG_PRINCIPAL, fg=FG_TEXTO, font=('Arial', 14)).pack(padx=30, pady=30)
        
        tk.Button(main_frame, text="Sí, acepto 👍", bg=FG_TITULO, fg='white', activebackground='#4da6ff',
                  font=('Arial', 14, 'bold'), command=self.create_personal_data_window, 
                  width=25, height=2, bd=0).pack(pady=10)
        
        tk.Button(main_frame, text="No deseo participar ❌", bg=BG_BOTON, fg='white', activebackground='#ff8c73',
                  font=('Arial', 14, 'bold'), command=self.root.destroy, 
                  width=25, height=2, bd=0).pack(pady=10)

    def create_personal_data_window(self):
        self.clear_window()
        data_frame = tk.Frame(self.root, bg=BG_ACENTO, padx=30, pady=30, relief='groove', bd=2)
        data_frame.pack(pady=50)

        tk.Label(data_frame, text="Datos Demográficos", font=("Arial", 18, 'bold'), bg=BG_ACENTO, fg=FG_TITULO).grid(row=0, column=0, pady=20, columnspan=2)
        labels = ["Nombre:", "Semestre:", "Edad:", "Correo electrónico:"]
        self.entries = []
        for idx, label in enumerate(labels):
            tk.Label(data_frame, text=label, bg=BG_ACENTO, fg=FG_TEXTO, font=('Arial', 12)).grid(row=idx+1, column=0, sticky='e', padx=15, pady=10)
            entry = tk.Entry(data_frame, width=40, font=('Arial', 12))
            entry.grid(row=idx+1, column=1, pady=10)
            self.entries.append(entry)

        tk.Button(data_frame, text="Siguiente ▶️", bg=BG_BOTON, fg='white', activebackground='#ff8c73',
                  font=('Arial', 14, 'bold'), command=self.save_personal_data, 
                  width=20, height=1, bd=0).grid(row=5, column=1, pady=30, sticky='e', padx=10)

    def save_personal_data(self):
        self.data['Nombre'] = self.entries[0].get()
        self.data['Semestre'] = self.entries[1].get()
        self.data['Edad'] = self.entries[2].get()
        self.data['Correo'] = self.entries[3].get()
        self.create_autoevaluacion_window()

    def create_autoevaluacion_window(self):
        self.clear_window()
        auto_frame = tk.Frame(self.root, bg=BG_ACENTO, padx=30, pady=30, relief='groove', bd=2)
        auto_frame.pack(pady=30)
        
        tk.Label(auto_frame, text="Autoevaluación (1-10)", font=("Arial", 18, 'bold'), bg=BG_ACENTO, fg=FG_TITULO).grid(row=0, column=0, pady=20, columnspan=2)
        instrucciones = "Te pedimos que a continuación califiques tus habilidades o conocimientos adquiridos sobre la tradición conductual, cognitivo conductual e interconductual del 1 al 10 en el que 1 es muy malos, y 10 es excelente."
        tk.Label(auto_frame, text=instrucciones, wraplength=700, justify='center', bg=BG_ACENTO, fg=FG_TEXTO, font=('Arial', 12)).grid(row=1, column=0, columnspan=2, padx=20, pady=15)

        self.auto_vars = {}
        secciones = ['Teórico', 'Metodológico', 'Aplicado']
        tradiciones = ['Conductual', 'Cognitivo Conductual', 'Interconductual']

        row = 2
        for sec in secciones:
            for trad in tradiciones:
                label = f"{sec} - {trad}"
                tk.Label(auto_frame, text=label, bg=BG_ACENTO, fg=FG_TEXTO, font=('Arial', 12)).grid(row=row, column=0, sticky='e', padx=15, pady=10)
                var = tk.IntVar()
                entry = tk.Entry(auto_frame, textvariable=var, width=15, font=('Arial', 12))
                entry.grid(row=row, column=1, pady=10)
                self.auto_vars[label] = var
                row += 1

        tk.Button(auto_frame, text="Siguiente ▶️", bg=BG_BOTON, fg='white', activebackground='#ff8c73',
                  font=('Arial', 14, 'bold'), command=self.save_autoevaluacion, 
                  width=20, height=1, bd=0).grid(row=row, column=1, pady=30, sticky='e', padx=10)

    def save_autoevaluacion(self):
        for key, var in self.auto_vars.items():
            try:
                valor = var.get()
                if 1 <= valor <= 10:
                    self.data[key] = valor
                else:
                    messagebox.showerror("Error de entrada", "Por favor califica tus habilidades con un número entre 1 y 10.")
                    return
            except:
                messagebox.showerror("Error de entrada", "Por favor ingresa solo números para la autoevaluación.")
                return
        self.create_questions_window()

    def create_questions_window(self):
        self.clear_window()
        self.current_question = 0
        self.correct_answers = 0
        
        # --- Listado de preguntas y opciones del Banco de Preguntas (Marzo 2026) ---
        base_questions = [
            {
                'pregunta': 'De los siguientes eventos históricos, elige aquel que fue crítico en la emergencia del conductismo.',
                'opciones': ['Teoría de la evolución de Darwin', 'Creación de la cámara de condicionamiento operante', 'Segunda guerra mundial', 'Derrumbamiento del muro de Berlín'],
                'correcta_text': 'Teoría de la evolución de Darwin',
                'categoria': 'Análisis conceptual de fenómenos psicológicos'
            },
            {
                'pregunta': 'En 1906 se publicó un libro que estableció los principios básicos de la acción refleja, incluyendo el concepto de sinapsis. ¿Cuál es el título de esa obra?',
                'opciones': ['Los reflejos condicionales de Iván Pávlov', 'La acción integrativa del sistema nervioso de Charles Sherrington.', 'La conducta de los organismos de B.F. Skinner.', 'Principios de Psicología de Keller y Schoenfeld.'],
                'correcta_text': 'La acción integrativa del sistema nervioso de Charles Sherrington.',
                'categoria': 'Análisis conceptual de fenómenos psicológicos'
            },
            {
                'pregunta': 'Thomas Hobbes sostenía que "todo lo que existe es materia; todo lo que ocurre es movimiento". ¿Qué postura filosófica representa esta afirmación, la cual anticipa el rechazo conductista de lo mental como sustancia independiente del cuerpo?',
                'opciones': ['Materialismo', 'Idealismo subjetivo', 'Dualismo interaccionista', 'Paralelismo psicofísico'],
                'correcta_text': 'Materialismo',
                'categoria': 'Análisis conceptual de fenómenos psicológicos'
            },
            {
                'pregunta': 'George Berkeley negó la existencia de la sustancia material y afirmó que la única realidad verdadera es la mente. ¿Qué postura filosófica defendía Berkeley y que es contraria a la del conductismo?',
                'opciones': ['Idealismo subjetivo', 'Materialismo', 'Dualismo interaccionista', 'Realismo ingenuo'],
                'correcta_text': 'Idealismo subjetivo',
                'categoria': 'Análisis conceptual de fenómenos psicológicos'
            },
            {
                'pregunta': 'Se refiere a la intensidad mínima de un estímulo para que pueda provocar una respuesta',
                'opciones': ['Umbral', 'Latencia', 'Tasa de respuesta', 'Intensidad'],
                'correcta_text': 'Umbral',
                'categoria': 'Análisis conceptual de fenómenos psicológicos'
            },
            {
                'pregunta': 'Es el estímulo cuya presencia señala que la respuesta no será reforzada',
                'opciones': ['Estímulo delta', 'Estímulo discriminativo', 'Estímulo condicional', 'Estímulo de extinción'],
                'correcta_text': 'Estímulo delta',
                'categoria': 'Análisis conceptual de fenómenos psicológicos'
            },
            {
                'pregunta': 'En un procedimiento de condicionamiento operante, es el estímulo que señala de manera consistente que una respuesta específica será seguida por una consecuencia reforzante',
                'opciones': ['Estímulo discriminativo', 'Estímulo delta', 'Estímulo incondicionado', 'Reforzador'],
                'correcta_text': 'Estímulo discriminativo',
                'categoria': 'Análisis conceptual de fenómenos psicológicos'
            },
            {
                'pregunta': 'El siguiente esquema es ejemplo del responder típico de un organismo al estar expuesto a un programa de reforzamiento denominado:',
                'opciones': ['Intervalo fijo', 'Reforzamiento operante', 'Intervalo variable', 'Razón fija'],
                'correcta_text': 'Intervalo fijo',
                'imagen_path': 'grafico_intervalo_fijo.png',
                'categoria': 'Análisis conceptual de fenómenos psicológicos'
            },
            {
                'pregunta': 'En un estudio clásico, la respuesta de miedo de un niño fue condicionada ante una rata blanca. Posteriormente, el niño manifestó reacciones similares frente a otros estímulos, como un conejo, un abrigo de piel y una máscara.\n\nLa figura siguiente presenta la intensidad de la respuesta emocional ante diversos estímulos, ordenados según su grado de semejanza con la rata blanca.\n\n¿Cuál de las siguientes opciones describe con mayor precisión la relación entre las propiedades de los estímulos y el patrón de respuesta observado?',
                'opciones': ['La similitud entre los estímulos determina la generalización de la respuesta, produciendo un gradiente decreciente.', 'La respuesta se mantiene constante, independientemente de las diferencias entre los estímulos.', 'La ausencia del estímulo aversivo elimina la respuesta de manera uniforme ante todos los estímulos.', 'La intensidad del estímulo aversivo explica por sí sola la variación en la respuesta.'],
                'correcta_text': 'La similitud entre los estímulos determina la generalización de la respuesta, produciendo un gradiente decreciente.',
                'imagen_path': 'gradiente_generalizacion.png',
                'categoria': 'Análisis conceptual de fenómenos psicológicos'
            },
            {
                'pregunta': 'El siguiente esquema es un ejemplo de un procedimiento de condicionamiento clásico denominado:',
                'opciones': ['Condicionamiento "huella"', 'Condicionamiento "para atrás"', 'Condicionamiento "simultáneo"', 'Condicionamiento pavloviano'],
                'correcta_text': 'Condicionamiento "huella"',
                'imagen_path': 'esquema_huella.png',
                'categoria': 'Análisis conceptual de fenómenos psicológicos'
            },
            {
                'pregunta': 'Es la unidad de análisis en la teoría del condicionamiento operante',
                'opciones': ['Triple relación de contingencia', 'estímulo–organismo–respuesta (E–O–R)', 'Tasa de respuesta', 'Contingencias de reforzamiento'],
                'correcta_text': 'Triple relación de contingencia',
                'categoria': 'Análisis metodológico de hechos psicológicos'
            },
            {
                'pregunta': 'Una rata es colocada en una cámara de condicionamiento operante. Cuando una luz se enciende, las respuestas de presión de la palanca son seguidas por la entrega de alimento. En ausencia de la luz, las presiones de palanca no producen ninguna consecuencia. El investigador busca analizar el comportamiento de la rata en función de estas condiciones. ¿Cuál de las siguientes opciones delimita correctamente la unidad de análisis pertinente para estudiar el fenómeno descrito?',
                'opciones': ['La relación entre la presencia de la luz, la presión de la palanca y la entrega de alimento', 'La frecuencia total de respuestas de presión de palanca durante la sesión', 'La presión de la palanca como conducta del organismo', 'El nivel de privación de alimento de la rata antes del experimento'],
                'correcta_text': 'La relación entre la presencia de la luz, la presión de la palanca y la entrega de alimento',
                'categoria': 'Análisis metodológico de hechos psicológicos'
            },
            {
                'pregunta': 'En los procedimientos de condicionamiento operante que utilizan laberintos, de las siguientes opciones ¿cuál suele ser una de las principales variables dependientes?',
                'opciones': ['Tiempo de recorrido', 'Reforzadores', 'Tasa de respuesta', 'Programa de reforzamiento'],
                'correcta_text': 'Tiempo de recorrido',
                'categoria': 'Análisis metodológico de hechos psicológicos'
            },
            {
                'pregunta': 'En un aula escolar, se te solicita registrar la conducta de los alumnos de levantarse de su asiento mientras la profesora imparte la clase. ¿cuál de estos registros consideras más apropiado para registrar la conducta de todos los alumnos?',
                'opciones': ['Registro Pla check', 'Registro acumulativo', 'Registro anecdótico', 'Registro frecuencia'],
                'correcta_text': 'Registro Pla check',
                'categoria': 'Análisis metodológico de hechos psicológicos'
            },
            {
                'pregunta': 'En un estudio experimental, se entrena a palomas para picar una tecla bajo distintos programas de reforzamiento. El investigador desea analizar cómo cambia la tasa de respuesta a lo largo de la sesión para identificar patrones característicos de cada programa. ¿Cuál de los siguientes registros es más adecuado para obtener datos que permitan satisfacer el objetivo escrito?',
                'opciones': ['Registro acumulativo continuo de respuestas', 'Muestreo por intervalos de tiempo fijo', 'Registro anecdótico de eventos relevantes', 'Registro de latencia de la primera respuesta'],
                'correcta_text': 'Registro acumulativo continuo de respuestas',
                'categoria': 'Análisis metodológico de hechos psicológicos'
            },
            {
                'pregunta': 'Cómo se denomina el procedimiento mediante el cual un organismo aprende una respuesta específica que y que consiste en el reforzamiento diferencial de aproximaciones sucesivas',
                'opciones': ['Moldeamiento', 'Modelamiento', 'Privación', 'Reforzamiento'],
                'correcta_text': 'Moldeamiento',
                'categoria': 'Análisis metodológico de hechos psicológicos'
            },
            {
                'pregunta': 'Dinsmoor y Lawson (1956) realizaron el siguiente experimento.\n\nObjetivo: Evaluar el efecto de la intensidad y el intervalo de tiempo que pospone un choque eléctrico sobre la latencia de la presión del operando de ratas en un programa de escape.\n\nCon base en la figura que se te presenta, elige la opción que mejor describa los resultados.',
                'opciones': ['A mayor intensidad y mayor tiempo sin choque eléctrico, menor latencia.', 'A mayor tiempo, mayor latencia.', 'Tiempo sin respuesta con choque eléctrico.', 'Mayor tiempo, menor latencia del choque eléctrico.'],
                'correcta_text': 'A mayor intensidad y mayor tiempo sin choque eléctrico, menor latencia.',
                'imagen_path': 'grafico_dinsmoor.png',
                'categoria': 'Análisis metodológico de hechos psicológicos'
            },
            {
                'pregunta': 'En el estudio Soares et al. (2025), se evaluó el efecto de la disminución en la magnitud del reforzamiento sobre la variabilidad conductual.\n\nEn la Figura 4 se presenta el índice de variabilidad de los sujetos bajo diferentes condiciones de magnitud del reforzador.\n\nCon base en los datos mostrados en la figura, selecciona la opción que describe correctamente el efecto de la disminución de la magnitud del reforzamiento sobre la variabilidad conductual.',
                'opciones': ['La disminución de la magnitud del reforzamiento se asocia con un incremento en la variabilidad conductual.', 'La variabilidad conductual disminuye conforme se reduce la magnitud del reforzamiento.', 'La variabilidad conductual se mantiene constante independientemente de la magnitud del reforzamiento.', 'La variabilidad conductual es mayor únicamente cuando la magnitud del reforzamiento es alta.'],
                'correcta_text': 'La disminución de la magnitud del reforzamiento se asocia con un incremento en la variabilidad conductual.',
                'imagen_path': 'grafica_latencia.png',
                'categoria': 'Análisis metodológico de hechos psicológicos'
            },
            {
                'pregunta': 'Briggs y Riccio (2007) estaban interesados en analizar si la extinción de una respuesta se veía afectada por la amnesia retrograda producida por eventos traumáticos. Para ello hicieron un estudio con seis grupos de ratas.\n\nUn primer grupo, denominado no ext, fue expuesto a situación en la que se condicionó la respuesta de evitación de ratas a trasladarse de una recámara blanca a otra de color negro en la que recibían descargas eléctricas.\n\nUn segundo grupo (ext) paso por la misma situación, pero tuvo una fase posterior en las que se les expuso en la sección en la que antes habían recibido choques eléctricos sin que estos se presentarán, es decir se extinguió su respuesta de evitación.\n\nUn tercer grupo (ext/hypo) paso por fase del grupo no ext, y una fase adicional en la que se les introdujo en agua fría hasta reducir considerablemente su temperatura corporal, y luego se introdujo de nuevo en la caja de vaivén, y se registró el tiempo que tardaban en introducirse en la cámara negra.\n\nTres grupos más, denominados 30, 33 y 37 pasaron por las mismas condiciones que el grupo ext/hypo pero, antes de ser reintroducidas a la caja de vaivén se les permitió llegar a los 30, 33 y 37 grados centígrados (respectivamente) y posteriormente se registró su comportamiento.\n\nLos datos recopilados se presentan en la siguiente figura, en la que se observa en el eje de las ordenadas el tiempo que tardaban en trasladarse de la sección blanca a la sección negra. Las barras de la abscisa son cada de uno de los grupos observados.\n\nCon base en los datos, elige aquella opción que represente un hallazgo reportado por los autores.',
                'opciones': ['La amnesia inducida por hipotermia interfiere con la recuperación de la extinción, y su efecto depende del estado fisiológico (temperatura corporal) en la reexposición.', 'La extinción elimina permanentemente la memoria del condicionamiento de evitación, independientemente del estado fisiológico del organismo.', 'La hipotermia incrementa de forma lineal la respuesta de evitación en todos los grupos experimentales sin interacción con la extinción.', 'El aprendizaje de evitación solo ocurre en ausencia de extinción y no puede recuperarse tras manipulación fisiológica posterior.'],
                'correcta_text': 'La amnesia inducida por hipotermia interfiere con la recuperación de la extinción, y su efecto depende del estado fisiológico (temperatura corporal) en la reexposición.',
                'imagen_path': 'briggs.png',
                'categoria': 'Análisis metodológico de hechos psicológicos'
            },
            {
                'pregunta': 'Antonio es un estudiante de Psicología que, tras haber tenido experiencias previas en las que las ratas de laboratorio se asociaron con sobresaltos intensos y situaciones desagradables durante prácticas (por ejemplo, movimientos bruscos de los animales y manipulación inesperada que le generaban una fuerte respuesta de miedo), desarrolla una reacción de temor al ver o estar cerca de ellas. Esta reacción incluye conductas de evitación como retirarse del lugar, cerrar los ojos o desviar la mirada.\n\nPara ayudarlo, su profesor lo expone de manera gradual y repetida a la presencia de ratas en un contexto controlado, en el que ahora las ratas no realizan movimientos bruscos ni ocurren situaciones desagradables. Antonio no puede evitar el estímulo, pero tampoco ocurre el evento aversivo que antes estaba asociado a las ratas. Con el paso de las sesiones, las respuestas de miedo de Antonio disminuyen hasta prácticamente desaparecer.\n\n¿Qué principio del aprendizaje explica este cambio en la conducta de Antonio?',
                'opciones': ['Extinción', 'Evitación condicional', 'Generalización', 'Reforzamiento'],
                'correcta_text': 'Extinción',
                'categoria': 'Intervención psicológica de problemas sociales'
            },
            {
                'pregunta': 'Adriana es una gerente bancaria de 45 años que hace ocho meses fue diagnosticada con diabetes. A partir del diagnóstico ha tenido tres crisis de salud asociadas a desajustes en su glucosa en sangre.\n\nCuando Adriana se encuentra en su rutina diaria y aparece la experiencia de antojo o deseo de consumir bebidas azucaradas, continúa consumiendo refresco de cola, rechaza la dieta prescrita argumentando que "no le sabe bien" y no toma el medicamento indicado por el médico. Como resultado inmediato, experimenta sensación de agrado al consumir el refresco y alimentos que prefiere, y además mantiene la creencia de que su condición puede controlarse sin medicamentos debido a la experiencia de su madre.\n\nCon base en lo anterior, determina cuál es la dimensión psicológica del caso.',
                'opciones': ['La interacción entre la experiencia de antojo, la respuesta de consumo de refresco y la sensación de agrado inmediata que mantiene la conducta.', 'Las recomendaciones médicas como antecedentes instruccionales que organizan el cumplimiento o incumplimiento del tratamiento.', 'La experiencia de antojo como evento antecedente que dispara la conducta de consumo.', 'La sensación de agrado inmediata como consecuencia que mantiene la conducta de consumo.'],
                'correcta_text': 'La interacción entre la experiencia de antojo, la respuesta de consumo de refresco y la sensación de agrado inmediata que mantiene la conducta.',
                'categoria': 'Intervención psicológica de problemas sociales'
            },
            {
                'pregunta': 'Luis es un estudiante universitario de 20 años que ha comenzado a reprobar varias materias. Refiere que cuando intenta estudiar, suele posponer la actividad revisando redes sociales o viendo videos en su celular. Aunque menciona sentirse preocupado por su desempeño académico, también comenta que "necesita relajarse un poco antes de empezar", lo que suele prolongarse durante varias horas. Luis reconoce que esta situación le ha generado conflictos familiares y académicos, pero le resulta difícil iniciar sus actividades escolares.\n\nCon base en lo anterior, determina cuál es la dimensión psicológica sobre la cual debe basarse el análisis psicológico.',
                'opciones': ['conducta de evitación de demandas académicas mantenido por consecuencias de alivio inmediato.', 'Déficit en habilidades de autorregulación cognitiva asociado a baja planificación.', 'Sesgo motivacional hacia reforzadores de alta inmediatez frente a metas de largo plazo.', 'Alteración en la percepción de autoeficacia académica ante tareas evaluativas.'],
                'correcta_text': 'conducta de evitación de demandas académicas mantenido por consecuencias de alivio inmediato.',
                'categoria': 'Intervención psicológica de problemas sociales'
            },
            {
                'pregunta': 'El dueño de una compañía de ventas te contrata para intervenir con un gerente de producto que, aunque cumple consistentemente con sus objetivos de ventas y desempeño, mantiene conflictos frecuentes con sus subordinados debido a que utiliza regaños, descalificaciones e insultos como forma habitual de comunicación. El objetivo de la intervención es reducir la frecuencia de estas conductas verbales aversivas sin afectar su nivel de desempeño laboral.\n\nAl entrevistarte con el gerente, ¿cuál de las siguientes preguntas consideras que te permitiría identificar la dimensión psicológica del problema?',
                'opciones': ['¿Qué situaciones ocurren en su jornada laboral en las que tiende a reaccionar con regaños o descalificaciones hacia sus subordinados, y qué cambia en su entorno inmediato después de hacerlo?', '¿Cómo interpreta usted el desempeño de sus subordinados y qué ideas tiene sobre la manera en que deberían responder a sus instrucciones dentro del equipo de trabajo?', '¿Qué aspectos de su formación profesional y experiencia previa considera que han influido en su estilo de comunicación actual con su equipo de trabajo?', '¿Qué cambios organizacionales, como metas de ventas o presión por resultados, han modificado la forma en que usted supervisa y evalúa el desempeño de su equipo?'],
                'correcta_text': '¿Qué situaciones ocurren en su jornada laboral en las que tiende a reaccionar con regaños o descalificaciones hacia sus subordinados, y qué cambia en su entorno inmediato después de hacerlo?',
                'categoria': 'Intervención psicológica de problemas sociales'
            },
            {
                'pregunta': 'Antonio es un niño de ocho años que acude a una escuela pública. Su profesora lo define como un niño inteligente y amable, pero que constantemente se distrae con sus compañeritos, lo que le impide terminar en tiempo y forma con las actividades que se le solicitan. Con el propósito de analizar las principales variables que afectan el termino de sus actividades, elige aquel registro que pueda proporcionar la información más apropiada.',
                'opciones': ['Registro de bloque temporal.', 'Registro anecdótico', 'Registro pla chek', 'Registro acumulativo'],
                'correcta_text': 'Registro de bloque temporal.',
                'categoria': 'Intervención psicológica de problemas sociales'
            },
            {
                'pregunta': 'Carmen es una estudiante de medicina que, en situaciones de nerviosismo, presenta una conducta repetitiva de rascado en el brazo, la cual le ha ocasionado lesiones visibles en la piel. Ha identificado que esta conducta ocurre con mayor intensidad durante la temporada de exámenes, particularmente ante la incertidumbre sobre su desempeño académico. Además, señala que el rascado tiende a ocurrir de manera continua o intermitente a lo largo de periodos prolongados, lo que dificulta delimitar con claridad el inicio y el final de cada episodio.\n\nDado este contexto: ¿Qué parámetro de registro conductual sería el más adecuado para evaluar esta conducta y por qué, considerando las características de ocurrencia descritas?',
                'opciones': ['Duración de la conducta, para cuantificar el tiempo total en que Carmen permanece rascándose.', 'Frecuencia de respuestas, para estimar cuántas veces ocurre la conducta en un periodo determinado.', 'Latencia de la respuesta, para identificar el tiempo que tarda en aparecer la conducta tras un estímulo.', 'Intervalos de muestreo, para estimar la ocurrencia de la conducta en periodos previamente definidos.'],
                'correcta_text': 'Duración de la conducta, para cuantificar el tiempo total en que Carmen permanece rascándose.',
                'categoria': 'Intervención psicológica de problemas sociales'
            },
            {
                'pregunta': "En un estudio experimental con 12 estudiantes universitarios con IMC en rango normopeso, se evaluó la regulación social de la conducta alimentaria mediante un diseño intrasujeto con cuatro fases (A, B, C y A'). En la fase A los participantes comían solos; en la fase B lo hacían con un modelo que consumía 113 g de alimentos poco calóricos en 1200 segundos; en la fase C con un modelo que consumía 133 g de alimentos calóricos en el mismo tiempo; y en la fase A' nuevamente solos. Se registró la cantidad de alimento ingerido, la frecuencia de elección de alimentos calóricos y la duración del consumo. Los resultados mostraron que la cantidad ingerida disminuyó en B (Mdn = 119.5 g) respecto a A (131 g) y aumentó en C (145 g), con diferencias significativas entre fases (χ²(3) = 12.67, p = .005). No hubo diferencias significativas en el tipo de alimento (p = .470). Los participantes con mayor consumo inicial mostraron mayor ajuste a las condiciones. Con base en esta información, ¿cuál es el análisis más adecuado de las variables implicadas?",
                'opciones': ['VI: consumo del modelo; VD: cantidad ingerida; moduladora: consumo inicial.', 'VI: hambre; VD: cantidad ingerida.', 'VI: tipo de alimento; VD: consumo.', 'VI: tiempo de sesión; VD: consumo.'],
                'correcta_text': 'VI: consumo del modelo; VD: cantidad ingerida; moduladora: consumo inicial.',
                'categoria': 'Intervención psicológica de problemas sociales'
            },
            {
                'pregunta': 'Armando es un peleador de muay thai que está por disputar un campeonato. En combates previos, cuando su oponente emite insultos dirigidos a su persona o a su familia, Armando incrementa abruptamente la frecuencia de ataques desorganizados, descuida la guardia y abandona la estrategia previamente acordada con su equipo, lo que lo expone a contraataques.\n\nSu próximo rival ha comenzado a utilizar sistemáticamente este tipo de provocaciones, por lo que se espera que estas ocurran durante la pelea. Aunque Armando identifica que estos estímulos le generan enojo intenso, su principal preocupación es que esta situación interfiera con su desempeño táctico.\n\nConsiderando que la intervención debe orientarse a modificar las condiciones asociadas a la problemática, ¿cuál de los siguientes objetivos es el más adecuado para guiar el diseño del procedimiento?',
                'opciones': ['Establecer que, ante la presencia de provocaciones verbales del oponente, Armando mantenga la secuencia táctica entrenada (guardia, distancia y combinación) en la mayoría de los intercambios durante el combate.', 'Entrenar a Armando para reducir la intensidad de su reacción ante los insultos mediante exposición progresiva a situaciones de provocación.', 'Lograr que Armando incremente la consistencia en la ejecución de combinaciones ofensivas durante la pelea.', 'Desarrollar en Armando la capacidad de ignorar los estímulos irrelevantes emitidos por el oponente durante el combate.'],
                'correcta_text': 'Establecer que, ante la presencia de provocaciones verbales del oponente, Armando mantenga la secuencia táctica entrenada (guardia, distancia y combinación) en la mayoría de los intercambios durante el combate.',
                'categoria': 'Intervención psicológica de problemas sociales'
            },
            {
                'pregunta': 'Carla trabaja en atención al cliente y reporta altos niveles de estrés laboral. Ante situaciones de presión, especialmente cuando interactúa con clientes conflictivos, tiende a evitar el contacto directo y delega la atención a sus compañeros. Esta conducta ha reducido momentáneamente su malestar, pero ha comenzado a generar problemas en su desempeño laboral y conflictos con su equipo.\n\nPara evaluar el progreso de una intervención dirigida a modificar esta problemática, ¿cuál de los siguientes indicadores es el más adecuado?',
                'opciones': ['Número de veces que Carla delega la atención de clientes conflictivos a sus compañeros por turno.', 'Nivel de estrés percibido por Carla al final de la jornada laboral.', 'Grado de satisfacción de Carla con su desempeño laboral.', 'Opinión de sus compañeros sobre la actitud de Carla en el trabaj'],
                'correcta_text': 'Número de veces que Carla delega la atención de clientes conflictivos a sus compañeros por turno.',
                'categoria': 'Intervención psicológica de problemas sociales'
            },
            {
                'pregunta': 'Diego es un estudiante de preparatoria que presenta dificultades para participar en clase. Aunque domina los contenidos, evita levantar la mano por miedo a equivocarse y ser juzgado por sus compañeros. Como consecuencia, su participación es baja y esto ha comenzado a afectar su calificación final. Se ha identificado que la conducta de evitación se mantiene porque reduce momentáneamente el malestar asociado a la evaluación social.\n\nCon base en esta situación, ¿cuál de los siguientes procedimientos es el más adecuado para intervenir el problema?',
                'opciones': ['Exponer gradualmente a Diego a situaciones de participación en clase, iniciando con contextos de baja exigencia y aumentando progresivamente la dificultad, reforzando sus intentos de participación.', 'Proporcionar información a Diego sobre la importancia de participar en clase para mejorar su rendimiento académico.', 'Solicitar a los compañeros que eviten emitir juicios negativos cuando Diego participe en clase.', 'Pedir a Diego que intente participar más frecuentemente durante las clases.'],
                'correcta_text': 'Exponer gradualmente a Diego a situaciones de participación en clase, iniciando con contextos de baja exigencia y aumentando progresivamente la dificultad, reforzando sus intentos de participación.',
                'categoria': 'Intervención psicológica de problemas sociales'
            },
            {
                'pregunta': 'Imagina que eres un psicólogo que trabaja en el ámbito educativo y recibes una solicitud para intervenir en la conducta de un niño que se levanta de su lugar de manera frecuente durante la clase. El objetivo es modificar las condiciones bajo las cuales esta conducta ocurre, de modo que el niño permanezca en su lugar y sólo se levante cuando la profesora lo indique. Con base en esta situación, ¿cuál es el principio psicológico más adecuado para orientar la intervención?',
                'opciones': ['Control de estímulos', 'Reforzamiento positivo', 'Castigo positivo', 'Extinción'],
                'correcta_text': 'Control de estímulos',
                'categoria': 'Intervención psicológica de problemas sociales'
            },
        ]
        
        # Procesamiento: Barajar opciones y reasignar el índice correcto
        self.questions = []
        for q in base_questions:
            correct_text = q['correcta_text']
            random.shuffle(q['opciones'])
            correct_index = q['opciones'].index(correct_text) + 1
            q['correcta'] = str(correct_index)
            q['Respuesta Correcta Texto'] = correct_text
            del q['correcta_text']
            self.questions.append(q)

        # Barajar el orden de las preguntas
        random.shuffle(self.questions)
        self.show_question()

    def show_question(self):
        self.clear_window()
        if self.current_question >= len(self.questions):
            self.show_results()
            return

        q = self.questions[self.current_question]
        
        has_image = 'imagen_path' in q
        # Tamaño de fuente homogéneo en todas las preguntas
        font_size_text = 12 
        
        # --- Configuración del Canvas y Scrollbar ---
        main_frame = tk.Frame(self.root, bg=BG_PRINCIPAL)
        main_frame.pack(pady=10, padx=50, fill='both', expand=True)

        canvas = tk.Canvas(main_frame, bg=BG_PRINCIPAL, highlightthickness=0)
        canvas.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollbar.pack(side="right", fill="y")

        canvas.configure(yscrollcommand=scrollbar.set)
        
        inner_frame = tk.Frame(canvas, bg=BG_ACENTO, padx=20, pady=20, relief='groove', bd=1)
        
        window_width = self.root.winfo_width()
        canvas.create_window((0, 0), window=inner_frame, anchor="nw", width=window_width if window_width > 1 else 900)

        inner_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion = canvas.bbox("all")))
        self.root.bind("<Configure>", lambda e: canvas.itemconfigure(canvas.winfo_children()[0], width=self.root.winfo_width() - 120))


        # --- Contenido de la Pregunta ---
        
        tk.Label(inner_frame, text=f"Pregunta {self.current_question + 1} de {len(self.questions)}", 
                 bg=BG_ACENTO, fg=FG_TITULO, font=('Arial', 14, 'italic')).pack(pady=5)
        
        for parrafo in q['pregunta'].split('\n\n'):
            tk.Label(inner_frame, text=parrafo.strip(), wraplength=800, justify='left',
                     bg=BG_ACENTO, fg=FG_TEXTO, font=('Arial', font_size_text)).pack(pady=(4,0), padx=10)

        self.current_image = None
        if has_image:
            try:
                img = Image.open(q['imagen_path'])
                width, height = img.size
                max_width = 450
                if width > max_width:
                    ratio = max_width / width
                    img = img.resize((max_width, int(height * ratio)), Image.LANCZOS)
                self.current_image = ImageTk.PhotoImage(img)
                tk.Label(inner_frame, image=self.current_image, bg=BG_ACENTO).pack(pady=15)
            except FileNotFoundError:
                tk.Label(inner_frame, text=f"[IMAGEN FALTANTE: {q['imagen_path']}]", fg='red', bg=BG_ACENTO, font=('Arial', 10)).pack(pady=5)
            except Exception:
                tk.Label(inner_frame, text=f"[ERROR DE CARGA DE IMAGEN: {q['imagen_path']}]", fg='red', bg=BG_ACENTO, font=('Arial', 10)).pack(pady=5)
        
        # Opciones de respuesta
        opciones = q['opciones']
        self.answer_var = tk.StringVar()
        
        options_frame = tk.Frame(inner_frame, bg=BG_ACENTO)
        options_frame.pack(pady=20, padx=10, fill='x')

        for idx, opt in enumerate(opciones, start=1):
            tk.Radiobutton(options_frame, text=f"{idx}) {opt}", variable=self.answer_var, value=str(idx), 
                           bg=BG_ACENTO, fg=FG_TEXTO, selectcolor=BG_PRINCIPAL,
                           wraplength=750, justify='left', font=('Arial', 12)).pack(anchor='w', pady=5)

        tk.Button(inner_frame, text="Siguiente ▶️", bg=BG_BOTON, fg='white', activebackground='#ff8c73',
                  font=('Arial', 14, 'bold'), command=self.check_answer, width=25, height=1, bd=0).pack(pady=30)
        
    def check_answer(self):
        q = self.questions[self.current_question]
        seleccion_index = self.answer_var.get()
        
        if not seleccion_index:
            messagebox.showwarning("Advertencia", "Por favor, selecciona una opción antes de continuar.")
            return

        opciones_ordenadas = q['opciones']
        respuesta_texto = opciones_ordenadas[int(seleccion_index) - 1]
        correcta = 'Correcta' if seleccion_index == q['correcta'] else 'Incorrecta'
        
        # Se guarda la respuesta en self.answers
        self.answers.append({
            'Pregunta': q['pregunta'],
            'Respuesta Seleccionada': respuesta_texto,
            'Respuesta Correcta Index': q['correcta'],
            'Respuesta Correcta Texto': q['Respuesta Correcta Texto'],
            'Resultado': correcta,
            'Categoria': q.get('categoria', 'N/A')
        })
        
        if seleccion_index == q['correcta']:
            self.correct_answers += 1
            
        self.current_question += 1
        self.show_question()

    def show_results(self):
        self.clear_window()
        
        num_questions = len(self.questions)
        promedio = (self.correct_answers / num_questions) * 100 if num_questions > 0 else 0

        self.data['Respuestas Correctas'] = self.correct_answers
        self.data['Num Preguntas'] = num_questions
        self.data['Promedio (%)'] = f"{promedio:.2f}"

        # Enviar datos a Google Sheets: una fila con todas las respuestas
        respuestas_lista = []
        for ans in self.answers:
            respuestas_lista.append({
                'pregunta': ans['Pregunta'],
                'respuesta': ans['Respuesta Seleccionada']
            })

        datos_sheets = {
            'fecha': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'nombre': self.data.get('Nombre', ''),
            'edad': self.data.get('Edad', ''),
            'semestre': self.data.get('Semestre', ''),
            'correo': self.data.get('Correo', ''),
            'respuestas': respuestas_lista,
        }
        enviar_a_sheets(datos_sheets)

        # Guardar y abrir el archivo Excel con el formato simplificado
        self.save_and_open_excel()
        
        tk.Label(self.root, text="¡Evaluación Finalizada! 🎉", font=('Arial', 24, 'bold'), fg=FG_TITULO, bg=BG_PRINCIPAL).pack(pady=40)
        tk.Label(self.root, text="Tus respuestas han sido guardadas. \nEl archivo de resultados se ha abierto. \nEso es todo, agradecemos tu participación. Por favor llama al investigador.", 
                 wraplength=700, justify='center', bg=BG_PRINCIPAL, fg=FG_TEXTO, font=('Arial', 14)).pack(pady=30)
        
        tk.Button(self.root, text="Cerrar Programa", bg=BG_BOTON, fg='white', activebackground='#ff8c73',
                  font=('Arial', 14, 'bold'), command=self.root.destroy, width=25, height=2, bd=0).pack(pady=40)

if __name__ == '__main__':
    root = tk.Tk()
    app = EvaluacionApp(root)
    root.mainloop()