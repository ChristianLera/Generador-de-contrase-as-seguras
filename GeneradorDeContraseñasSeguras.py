import secrets
import string
import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime
import os
import pyperclip
from pathlib import Path

class GestorContrasenas:
    """
    Gestor de contraseñas con almacenamiento en Excel
    """
    
    def __init__(self):
        self.archivo_excel = "mis_contrasenas.xlsx"
        self.df = None
        self.cargar_datos()
        
    def cargar_datos(self):
        """Carga los datos del archivo Excel o crea uno nuevo"""
        if os.path.exists(self.archivo_excel):
            try:
                self.df = pd.read_excel(self.archivo_excel)
                # Asegurar que todas las columnas existan
                columnas_requeridas = ['ID', 'Fecha', 'Descripción', 'Contraseña', 
                                      'Longitud', 'Entropía', 'Nivel_Seguridad', 'Categoría']
                for col in columnas_requeridas:
                    if col not in self.df.columns:
                        self.df[col] = ''
            except:
                self.crear_nuevo_archivo()
        else:
            self.crear_nuevo_archivo()
            
    def crear_nuevo_archivo(self):
        """Crea un nuevo archivo Excel con la estructura base"""
        self.df = pd.DataFrame(columns=['ID', 'Fecha', 'Descripción', 'Contraseña', 
                                        'Longitud', 'Entropía', 'Nivel_Seguridad', 'Categoría'])
        self.guardar()
        
    def guardar(self):
        """Guarda los datos en el archivo Excel"""
        try:
            self.df.to_excel(self.archivo_excel, index=False, engine='openpyxl')
            return True
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar: {str(e)}")
            return False
            
    def agregar_contrasena(self, descripcion, contrasena, entropia, nivel, categoria="General"):
        """Agrega una nueva contraseña al registro"""
        nuevo_id = len(self.df) + 1
        nueva_fila = {
            'ID': nuevo_id,
            'Fecha': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'Descripción': descripcion,
            'Contraseña': contrasena,
            'Longitud': len(contrasena),
            'Entropía': entropia,
            'Nivel_Seguridad': nivel,
            'Categoría': categoria
        }
        self.df = pd.concat([self.df, pd.DataFrame([nueva_fila])], ignore_index=True)
        return self.guardar()
    
    def obtener_todas(self):
        """Obtiene todas las contraseñas ordenadas por fecha descendente"""
        if self.df is not None and not self.df.empty:
            return self.df.sort_values('Fecha', ascending=False)
        return pd.DataFrame()
    
    def eliminar_contrasena(self, id_contrasena):
        """Elimina una contraseña por ID"""
        self.df = self.df[self.df['ID'] != id_contrasena]
        # Reorganizar IDs
        self.df['ID'] = range(1, len(self.df) + 1)
        return self.guardar()
    
    def buscar_por_descripcion(self, texto):
        """Busca contraseñas por descripción"""
        if self.df is not None and not self.df.empty:
            mascara = self.df['Descripción'].str.contains(texto, case=False, na=False)
            return self.df[mascara].sort_values('Fecha', ascending=False)
        return pd.DataFrame()
    
    def exportar_csv(self, archivo):
        """Exporta los datos a CSV"""
        try:
            self.df.to_csv(archivo, index=False, encoding='utf-8-sig')
            return True
        except:
            return False
            
    def importar_csv(self, archivo):
        """Importa datos desde CSV"""
        try:
            df_importado = pd.read_csv(archivo, encoding='utf-8-sig')
            # Verificar columnas necesarias
            columnas_necesarias = ['Descripción', 'Contraseña', 'Categoría']
            if all(col in df_importado.columns for col in columnas_necesarias):
                for _, row in df_importado.iterrows():
                    self.agregar_contrasena(
                        row['Descripción'],
                        row['Contraseña'],
                        "N/A",
                        "Importada",
                        row.get('Categoría', 'Importada')
                    )
                return True
            return False
        except:
            return False


class GeneradorContrasenasGUI:
    """
    Aplicación principal con interfaz gráfica
    """
    
    def __init__(self):
        self.ventana = tk.Tk()
        self.ventana.title("🔐 Gestor Profesional de Contraseñas Seguras")
        self.ventana.geometry("900x700")
        self.ventana.resizable(True, True)
        
        # Inicializar gestor
        self.gestor = GestorContrasenas()
        
        # Configurar estilo
        self.estilo = ttk.Style()
        self.estilo.theme_use('clam')
        
        self.colores = {
            'bg': '#1e1e2e',
            'fg': '#cdd6f4',
            'acento': '#89b4fa',
            'exito': '#a6e3a1',
            'peligro': '#f38ba8',
            'advertencia': '#f9e2af'
        }
        
        self.ventana.configure(bg=self.colores['bg'])
        
        # Caracteres disponibles
        self.MAYUSCULAS = string.ascii_uppercase
        self.MINUSCULAS = string.ascii_lowercase
        self.DIGITOS = string.digits
        self.SIMBOLOS = "!@#$%^&*()_+-=[]{}|;:,.<>?"
        self.AMBIGUOS = "il1Lo0O"
        
        # Variables
        self.contrasena_actual = tk.StringVar()
        self.descripcion_actual = tk.StringVar()
        self.categoria_actual = tk.StringVar(value="General")
        
        # Crear interfaz
        self._crear_interfaz()
        
        # Cargar datos iniciales
        self._actualizar_lista_contrasenas()
        
    def _crear_interfaz(self):
        """Crea todos los elementos de la interfaz"""
        
        # Frame principal
        main_frame = ttk.Frame(self.ventana)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Notebook principal
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # Pestañas
        self._crear_pestana_generador(notebook)
        self._crear_pestana_boveda(notebook)
        self._crear_pestana_estadisticas(notebook)
        self._crear_pestana_configuracion(notebook)
        
    def _crear_pestana_generador(self, parent):
        """Pestaña del generador de contraseñas"""
        frame = ttk.Frame(parent)
        parent.add(frame, text="🔐 Generador")
        
        # Frame de contraseña generada
        frame_pass = tk.Frame(frame, bg=self.colores['bg'])
        frame_pass.pack(fill=tk.X, padx=20, pady=20)
        
        tk.Label(
            frame_pass,
            text="Contraseña Generada:",
            font=('Arial', 12, 'bold'),
            bg=self.colores['bg'],
            fg=self.colores['fg']
        ).pack(anchor=tk.W)
        
        # Entry para mostrar contraseña
        entry_pass = tk.Entry(
            frame_pass,
            textvariable=self.contrasena_actual,
            font=('Courier', 14),
            bg='#2d2d3d',
            fg=self.colores['exito'],
            state='readonly',
            relief=tk.SUNKEN
        )
        entry_pass.pack(fill=tk.X, pady=10, ipady=8)
        
        # Botones de acción
        frame_botones = tk.Frame(frame_pass, bg=self.colores['bg'])
        frame_botones.pack(fill=tk.X, pady=5)
        
        tk.Button(
            frame_botones,
            text="🔄 Generar",
            command=self._generar_contrasena,
            bg=self.colores['exito'],
            fg='black',
            font=('Arial', 10, 'bold'),
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            frame_botones,
            text="📋 Copiar",
            command=self._copiar_portapapeles,
            bg=self.colores['acento'],
            fg='white',
            font=('Arial', 10, 'bold'),
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
        
        # Frame de configuración
        frame_config = tk.LabelFrame(
            frame,
            text="Configuración",
            bg=self.colores['bg'],
            fg=self.colores['fg'],
            font=('Arial', 11, 'bold')
        )
        frame_config.pack(fill=tk.X, padx=20, pady=10)
        
        # Longitud
        tk.Label(
            frame_config,
            text="Longitud:",
            bg=self.colores['bg'],
            fg=self.colores['fg']
        ).grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        
        self.longitud_var = tk.IntVar(value=20)
        longitud_scale = tk.Scale(
            frame_config,
            from_=8,
            to=64,
            orient=tk.HORIZONTAL,
            variable=self.longitud_var,
            bg=self.colores['bg'],
            fg=self.colores['fg'],
            highlightthickness=0,
            length=200
        )
        longitud_scale.grid(row=0, column=1, padx=10, pady=5)
        
        # Tipos de caracteres
        self.incluir_mayus = tk.BooleanVar(value=True)
        self.incluir_minus = tk.BooleanVar(value=True)
        self.incluir_digitos = tk.BooleanVar(value=True)
        self.incluir_simbolos = tk.BooleanVar(value=True)
        self.evitar_ambiguos = tk.BooleanVar(value=True)
        
        checks = [
            ("Mayúsculas", self.incluir_mayus),
            ("Minúsculas", self.incluir_minus),
            ("Dígitos", self.incluir_digitos),
            ("Símbolos", self.incluir_simbolos),
            ("Evitar ambiguos", self.evitar_ambiguos)
        ]
        
        for i, (texto, var) in enumerate(checks):
            tk.Checkbutton(
                frame_config,
                text=texto,
                variable=var,
                bg=self.colores['bg'],
                fg=self.colores['fg'],
                selectcolor=self.colores['bg']
            ).grid(row=i+1, column=0, sticky=tk.W, padx=10, pady=2)
        
        # Frame para guardar
        frame_guardar = tk.LabelFrame(
            frame,
            text="Guardar Contraseña",
            bg=self.colores['bg'],
            fg=self.colores['fg'],
            font=('Arial', 11, 'bold')
        )
        frame_guardar.pack(fill=tk.X, padx=20, pady=10)
        
        tk.Label(
            frame_guardar,
            text="Descripción:",
            bg=self.colores['bg'],
            fg=self.colores['fg']
        ).grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        
        tk.Entry(
            frame_guardar,
            textvariable=self.descripcion_actual,
            font=('Arial', 10),
            bg='#2d2d3d',
            fg=self.colores['fg'],
            width=40
        ).grid(row=0, column=1, padx=10, pady=5, columnspan=2)
        
        tk.Label(
            frame_guardar,
            text="Categoría:",
            bg=self.colores['bg'],
            fg=self.colores['fg']
        ).grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        
        categorias = ["General", "Email", "Bancos", "Redes Sociales", "Trabajo", "Personal", "Otros"]
        self.categoria_combo = ttk.Combobox(
            frame_guardar,
            textvariable=self.categoria_actual,
            values=categorias,
            width=37
        )
        self.categoria_combo.grid(row=1, column=1, padx=10, pady=5, columnspan=2)
        
        tk.Button(
            frame_guardar,
            text="💾 Guardar Contraseña",
            command=self._guardar_contrasena_actual,
            bg=self.colores['advertencia'],
            fg='black',
            font=('Arial', 11, 'bold'),
            cursor='hand2'
        ).grid(row=2, column=0, columnspan=3, pady=10)
        
        # Generar primera contraseña
        self._generar_contrasena()
        
    def _crear_pestana_boveda(self, parent):
        """Pestaña de bóveda de contraseñas"""
        frame = ttk.Frame(parent)
        parent.add(frame, text="📦 Bóveda de Contraseñas")
        
        # Frame de búsqueda
        frame_busqueda = tk.Frame(frame, bg=self.colores['bg'])
        frame_busqueda.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Label(
            frame_busqueda,
            text="Buscar:",
            bg=self.colores['bg'],
            fg=self.colores['fg']
        ).pack(side=tk.LEFT, padx=5)
        
        self.busqueda_var = tk.StringVar()
        self.busqueda_var.trace('w', lambda *args: self._buscar_contrasenas())
        
        tk.Entry(
            frame_busqueda,
            textvariable=self.busqueda_var,
            font=('Arial', 10),
            bg='#2d2d3d',
            fg=self.colores['fg'],
            width=40
        ).pack(side=tk.LEFT, padx=5)
        
        # Treeview para mostrar contraseñas
        frame_tree = tk.Frame(frame, bg=self.colores['bg'])
        frame_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Scrollbars
        scroll_y = tk.Scrollbar(frame_tree, orient=tk.VERTICAL)
        scroll_x = tk.Scrollbar(frame_tree, orient=tk.HORIZONTAL)
        
        self.tree = ttk.Treeview(
            frame_tree,
            columns=('ID', 'Fecha', 'Descripción', 'Contraseña', 'Longitud', 'Nivel', 'Categoría'),
            show='headings',
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set
        )
        
        scroll_y.config(command=self.tree.yview)
        scroll_x.config(command=self.tree.xview)
        
        # Configurar columnas
        self.tree.heading('ID', text='ID')
        self.tree.heading('Fecha', text='Fecha')
        self.tree.heading('Descripción', text='Descripción')
        self.tree.heading('Contraseña', text='Contraseña')
        self.tree.heading('Longitud', text='Longitud')
        self.tree.heading('Nivel', text='Seguridad')
        self.tree.heading('Categoría', text='Categoría')
        
        self.tree.column('ID', width=50)
        self.tree.column('Fecha', width=150)
        self.tree.column('Descripción', width=200)
        self.tree.column('Contraseña', width=200)
        self.tree.column('Longitud', width=70)
        self.tree.column('Nivel', width=100)
        self.tree.column('Categoría', width=120)
        
        self.tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Botones de acción
        frame_acciones = tk.Frame(frame, bg=self.colores['bg'])
        frame_acciones.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Button(
            frame_acciones,
            text="📋 Copiar Contraseña",
            command=self._copiar_seleccionada,
            bg=self.colores['acento'],
            fg='white',
            font=('Arial', 10, 'bold'),
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            frame_acciones,
            text="🗑️ Eliminar Seleccionada",
            command=self._eliminar_seleccionada,
            bg=self.colores['peligro'],
            fg='white',
            font=('Arial', 10, 'bold'),
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            frame_acciones,
            text="📤 Exportar CSV",
            command=self._exportar_csv,
            bg=self.colores['exito'],
            fg='black',
            font=('Arial', 10, 'bold'),
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            frame_acciones,
            text="📥 Importar CSV",
            command=self._importar_csv,
            bg=self.colores['advertencia'],
            fg='black',
            font=('Arial', 10, 'bold'),
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
        
    def _crear_pestana_estadisticas(self, parent):
        """Pestaña de estadísticas"""
        frame = ttk.Frame(parent)
        parent.add(frame, text="📊 Estadísticas")
        
        # Text widget para estadísticas
        self.estadisticas_text = tk.Text(
            frame,
            bg='#2d2d3d',
            fg=self.colores['fg'],
            font=('Consolas', 11),
            wrap=tk.WORD
        )
        self.estadisticas_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        tk.Button(
            frame,
            text="🔄 Actualizar Estadísticas",
            command=self._actualizar_estadisticas,
            bg=self.colores['acento'],
            fg='white',
            font=('Arial', 10, 'bold'),
            cursor='hand2'
        ).pack(pady=10)
        
        self._actualizar_estadisticas()
        
    def _crear_pestana_configuracion(self, parent):
        """Pestaña de configuración"""
        frame = ttk.Frame(parent)
        parent.add(frame, text="⚙️ Configuración")
        
        # Configuración del archivo
        frame_archivo = tk.LabelFrame(
            frame,
            text="Archivo de Datos",
            bg=self.colores['bg'],
            fg=self.colores['fg'],
            font=('Arial', 11, 'bold')
        )
        frame_archivo.pack(fill=tk.X, padx=20, pady=20)
        
        tk.Label(
            frame_archivo,
            text=f"Archivo actual: {self.gestor.archivo_excel}",
            bg=self.colores['bg'],
            fg=self.colores['fg']
        ).pack(pady=10)
        
        tk.Button(
            frame_archivo,
            text="📁 Cambiar ubicación",
            command=self._cambiar_ubicacion,
            bg=self.colores['advertencia'],
            fg='black',
            font=('Arial', 10, 'bold'),
            cursor='hand2'
        ).pack(pady=5)
        
        # Botón de respaldo
        tk.Button(
            frame_archivo,
            text="💾 Crear respaldo",
            command=self._crear_respaldo,
            bg=self.colores['exito'],
            fg='black',
            font=('Arial', 10, 'bold'),
            cursor='hand2'
        ).pack(pady=5)
        
    def _generar_contrasena(self):
        """Genera una nueva contraseña"""
        longitud = self.longitud_var.get()
        
        charset = ""
        if self.incluir_mayus.get():
            charset += self.MAYUSCULAS
        if self.incluir_minus.get():
            charset += self.MINUSCULAS
        if self.incluir_digitos.get():
            charset += self.DIGITOS
        if self.incluir_simbolos.get():
            charset += self.SIMBOLOS
            
        if self.evitar_ambiguos.get():
            for char in self.AMBIGUOS:
                charset = charset.replace(char, '')
        
        if not charset:
            messagebox.showerror("Error", "Selecciona al menos un tipo de carácter")
            return
        
        for _ in range(10):
            password = ''.join(secrets.choice(charset) for _ in range(longitud))
            
            cumple = True
            if self.incluir_mayus.get() and not any(c in self.MAYUSCULAS for c in password):
                cumple = False
            if self.incluir_minus.get() and not any(c in self.MINUSCULAS for c in password):
                cumple = False
            if self.incluir_digitos.get() and not any(c in self.DIGITOS for c in password):
                cumple = False
            if self.incluir_simbolos.get() and not any(c in self.SIMBOLOS for c in password):
                cumple = False
                
            if cumple:
                self.contrasena_actual.set(password)
                return
        
        messagebox.showerror("Error", "No se pudo generar la contraseña")
        
    def _guardar_contrasena_actual(self):
        """Guarda la contraseña actual en la bóveda"""
        contrasena = self.contrasena_actual.get()
        descripcion = self.descripcion_actual.get().strip()
        
        if not contrasena:
            messagebox.showwarning("Advertencia", "No hay contraseña para guardar")
            return
            
        if not descripcion:
            messagebox.showwarning("Advertencia", "Por favor, ingresa una descripción")
            return
            
        # Calcular entropía y nivel
        entropia = self._calcular_entropia(contrasena)
        nivel = self._calcular_nivel(contrasena)
        
        if self.gestor.agregar_contrasena(descripcion, contrasena, entropia, nivel, self.categoria_actual.get()):
            messagebox.showinfo("Éxito", "¡Contraseña guardada correctamente!")
            self.descripcion_actual.set("")
            self._actualizar_lista_contrasenas()
            self._actualizar_estadisticas()
        else:
            messagebox.showerror("Error", "No se pudo guardar la contraseña")
            
    def _actualizar_lista_contrasenas(self, busqueda=None):
        """Actualiza el treeview con las contraseñas"""
        # Limpiar treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Obtener datos
        if busqueda:
            df = self.gestor.buscar_por_descripcion(busqueda)
        else:
            df = self.gestor.obtener_todas()
            
        if df.empty:
            return
            
        # Agregar filas
        for _, row in df.iterrows():
            # Enmascarar contraseña por seguridad
            password = row['Contraseña']
            if len(password) > 10:
                password_display = password[:5] + "..." + password[-3:]
            else:
                password_display = "*" * len(password)
                
            self.tree.insert('', 'end', values=(
                row['ID'],
                row['Fecha'],
                row['Descripción'],
                password_display,
                row['Longitud'],
                row['Nivel_Seguridad'],
                row['Categoría']
            ), tags=('real_pass', row['Contraseña']))
            
    def _buscar_contrasenas(self):
        """Busca contraseñas por descripción"""
        texto = self.busqueda_var.get().strip()
        if texto:
            self._actualizar_lista_contrasenas(texto)
        else:
            self._actualizar_lista_contrasenas()
            
    def _copiar_seleccionada(self):
        """Copia la contraseña seleccionada al portapapeles"""
        seleccion = self.tree.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Selecciona una contraseña primero")
            return
            
        # Obtener la contraseña real del tag
        item = self.tree.item(seleccion[0])
        password_real = item['tags'][1] if item['tags'] else None
        
        if password_real:
            pyperclip.copy(password_real)
            messagebox.showinfo("Éxito", "¡Contraseña copiada al portapapeles!")
            
    def _eliminar_seleccionada(self):
        """Elimina la contraseña seleccionada"""
        seleccion = self.tree.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Selecciona una contraseña primero")
            return
            
        if messagebox.askyesno("Confirmar", "¿Estás seguro de eliminar esta contraseña?"):
            item = self.tree.item(seleccion[0])
            id_contrasena = item['values'][0]
            
            if self.gestor.eliminar_contrasena(id_contrasena):
                messagebox.showinfo("Éxito", "Contraseña eliminada")
                self._actualizar_lista_contrasenas()
                self._actualizar_estadisticas()
                
    def _copiar_portapapeles(self):
        """Copia la contraseña actual"""
        contrasena = self.contrasena_actual.get()
        if contrasena:
            pyperclip.copy(contrasena)
            messagebox.showinfo("Éxito", "¡Contraseña copiada!")
            
    def _exportar_csv(self):
        """Exporta datos a CSV"""
        archivo = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if archivo:
            if self.gestor.exportar_csv(archivo):
                messagebox.showinfo("Éxito", f"Datos exportados a {archivo}")
            else:
                messagebox.showerror("Error", "No se pudo exportar")
                
    def _importar_csv(self):
        """Importa datos desde CSV"""
        archivo = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if archivo:
            if self.gestor.importar_csv(archivo):
                messagebox.showinfo("Éxito", "Datos importados correctamente")
                self._actualizar_lista_contrasenas()
                self._actualizar_estadisticas()
            else:
                messagebox.showerror("Error", "No se pudo importar. Verifica el formato")
                
    def _actualizar_estadisticas(self):
        """Actualiza la pestaña de estadísticas"""
        df = self.gestor.obtener_todas()
        
        if df.empty:
            self.estadisticas_text.delete(1.0, tk.END)
            self.estadisticas_text.insert(1.0, "No hay contraseñas guardadas aún.")
            return
            
        stats = f"""
{'═' * 60}
📊 ESTADÍSTICAS DE TU BÓVEDA
{'═' * 60}

📈 RESÚMEN GENERAL:
• Total de contraseñas: {len(df)}
• Longitud promedio: {df['Longitud'].mean():.1f} caracteres
• Contraseñas por categoría:

"""
        
        # Contar por categoría
        categorias = df['Categoría'].value_counts()
        for cat, count in categorias.items():
            stats += f"  • {cat}: {count}\n"
            
        stats += f"""
{'─' * 60}
🔒 NIVEL DE SEGURIDAD:
"""
        
        niveles = df['Nivel_Seguridad'].value_counts()
        for nivel, count in niveles.items():
            stats += f"  • {nivel}: {count}\n"
            
        stats += f"""
{'─' * 60}
📅 ÚLTIMAS CONTRASEÑAS AGREGADAS:
"""
        
        ultimas = df.head(5)
        for _, row in ultimas.iterrows():
            stats += f"  • {row['Fecha']} - {row['Descripción']} ({row['Categoría']})\n"
            
        self.estadisticas_text.delete(1.0, tk.END)
        self.estadisticas_text.insert(1.0, stats)
        
    def _calcular_entropia(self, password):
        """Calcula entropía aproximada"""
        charset_size = len(set(password))
        return len(password) * (charset_size.bit_length() - 1)
        
    def _calcular_nivel(self, password):
        """Calcula nivel de seguridad"""
        puntuacion = 0
        if any(c in self.MAYUSCULAS for c in password): puntuacion += 15
        if any(c in self.MINUSCULAS for c in password): puntuacion += 15
        if any(c in self.DIGITOS for c in password): puntuacion += 20
        if any(c in self.SIMBOLOS for c in password): puntuacion += 25
        
        if len(password) >= 20:
            puntuacion += 15
        elif len(password) >= 16:
            puntuacion += 10
        elif len(password) >= 12:
            puntuacion += 5
            
        if puntuacion >= 80:
            return "Excelente"
        elif puntuacion >= 60:
            return "Buena"
        else:
            return "Débil"
            
    def _cambiar_ubicacion(self):
        """Cambia la ubicación del archivo Excel"""
        archivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if archivo:
            # Copiar datos al nuevo archivo
            nuevo_gestor = GestorContrasenas()
            nuevo_gestor.archivo_excel = archivo
            nuevo_gestor.df = self.gestor.df.copy()
            if nuevo_gestor.guardar():
                self.gestor = nuevo_gestor
                messagebox.showinfo("Éxito", f"Base de datos movida a {archivo}")
                
    def _crear_respaldo(self):
        """Crea un respaldo del archivo actual"""
        if os.path.exists(self.gestor.archivo_excel):
            backup_name = f"respaldo_contrasenas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            import shutil
            shutil.copy(self.gestor.archivo_excel, backup_name)
            messagebox.showinfo("Éxito", f"Respaldo creado: {backup_name}")
            
    def ejecutar(self):
        """Inicia la aplicación"""
        self.ventana.mainloop()


if __name__ == "__main__":
    # Verificar dependencias
    try:
        import pandas
        import openpyxl
        import pyperclip
    except ImportError:
        import subprocess
        print("Instalando dependencias necesarias...")
        subprocess.check_call(['pip', 'install', 'pandas', 'openpyxl', 'pyperclip'])
        
    app = GeneradorContrasenasGUI()
    app.ejecutar()
