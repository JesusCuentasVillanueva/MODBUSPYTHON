import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import time
import threading
from pymodbus.client import ModbusSerialClient
import logging
import json
import os
from datetime import datetime

# Configurar logging
logging.basicConfig()
log = logging.getLogger()
log.setLevel(logging.INFO)

class ModbusSlaveFinderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Buscador de Esclavos Modbus")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # Variables para la configuración
        self.port_var = tk.StringVar(value="COM3")
        self.baudrate_var = tk.IntVar(value=9600)
        self.parity_var = tk.StringVar(value="N")
        self.stopbits_var = tk.IntVar(value=1)
        self.bytesize_var = tk.IntVar(value=8)
        self.timeout_var = tk.DoubleVar(value=0.1)  # Timeout más corto para escaneo rápido
        
        # Variables para la búsqueda
        self.slave_start_var = tk.IntVar(value=1)
        self.slave_end_var = tk.IntVar(value=247)  # Máximo ID para Modbus
        self.slave_test_function_var = tk.StringVar(value="holding")
        self.slave_test_register_var = tk.IntVar(value=0)
        
        # Variables para el control de la búsqueda
        self.slave_finding = False
        self.slave_finder_thread = None
        self.slave_finder_client = None
        
        # Crear interfaz
        self.create_widgets()
        
        # Configurar cierre de la aplicación
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def create_widgets(self):
        # Panel superior: Configuración de comunicación
        config_frame = ttk.LabelFrame(self.root, text="Configuración de comunicación", padding=10)
        config_frame.pack(fill=tk.X, expand=False, pady=(10, 5), padx=10)
        
        # Fila 1
        frame1 = ttk.Frame(config_frame)
        frame1.pack(fill=tk.X, pady=5)
        
        ttk.Label(frame1, text="Puerto:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(frame1, textvariable=self.port_var, width=10).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(frame1, text="Baudrate:").pack(side=tk.LEFT, padx=5)
        ttk.Combobox(frame1, textvariable=self.baudrate_var, 
                    values=[1200, 2400, 4800, 9600, 19200, 38400, 57600, 115200], 
                    width=10).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(frame1, text="Paridad:").pack(side=tk.LEFT, padx=5)
        ttk.Combobox(frame1, textvariable=self.parity_var, 
                    values=["N", "E", "O"], width=5).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(frame1, text="Bits de parada:").pack(side=tk.LEFT, padx=5)
        ttk.Combobox(frame1, textvariable=self.stopbits_var, 
                    values=[1, 2], width=5).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(frame1, text="Bits de datos:").pack(side=tk.LEFT, padx=5)
        ttk.Combobox(frame1, textvariable=self.bytesize_var, 
                    values=[7, 8], width=5).pack(side=tk.LEFT, padx=5)
        
        # Panel de configuración de búsqueda
        search_frame = ttk.LabelFrame(self.root, text="Configuración de búsqueda", padding=10)
        search_frame.pack(fill=tk.X, expand=False, pady=5, padx=10)
        
        # Fila 1
        frame2 = ttk.Frame(search_frame)
        frame2.pack(fill=tk.X, pady=5)
        
        ttk.Label(frame2, text="ID inicial:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(frame2, textvariable=self.slave_start_var, width=8).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(frame2, text="ID final:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(frame2, textvariable=self.slave_end_var, width=8).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(frame2, text="Función de prueba:").pack(side=tk.LEFT, padx=5)
        ttk.Combobox(frame2, textvariable=self.slave_test_function_var, 
                    values=["holding", "input", "coil", "discrete_input"], width=13).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(frame2, text="Registro de prueba:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(frame2, textvariable=self.slave_test_register_var, width=8).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(frame2, text="Timeout (s):").pack(side=tk.LEFT, padx=5)
        ttk.Entry(frame2, textvariable=self.timeout_var, width=8).pack(side=tk.LEFT, padx=5)
        
        # Botones de control
        button_frame = ttk.Frame(search_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        self.start_button = ttk.Button(button_frame, text="Iniciar búsqueda", 
                                     command=self.start_slave_finder)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(button_frame, text="Detener", 
                                    command=self.stop_slave_finder, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        # Barra de progreso
        self.progress_frame = ttk.Frame(search_frame)
        self.progress_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(self.progress_frame, text="Progreso:").pack(side=tk.LEFT, padx=5)
        self.progress_var = tk.StringVar(value="0 / 0")
        ttk.Label(self.progress_frame, textvariable=self.progress_var).pack(side=tk.LEFT, padx=5)
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, orient=tk.HORIZONTAL, length=100, mode='determinate')
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Panel de resultados
        results_frame = ttk.LabelFrame(self.root, text="Esclavos encontrados", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True, pady=5, padx=10)
        
        # Tabla para los resultados
        columns = ("slave_id", "response_time", "register_value", "status")
        self.results_tree = ttk.Treeview(results_frame, columns=columns, show="headings")
        self.results_tree.heading("slave_id", text="ID Esclavo")
        self.results_tree.heading("response_time", text="Tiempo de respuesta (ms)")
        self.results_tree.heading("register_value", text="Valor del registro")
        self.results_tree.heading("status", text="Estado")
        self.results_tree.column("slave_id", width=100, anchor=tk.CENTER)
        self.results_tree.column("response_time", width=150, anchor=tk.CENTER)
        self.results_tree.column("register_value", width=150, anchor=tk.CENTER)
        self.results_tree.column("status", width=150, anchor=tk.CENTER)
        self.results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Scrollbar para la tabla
        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.results_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.results_tree.configure(yscrollcommand=scrollbar.set)
        
        # Botones para exportar datos
        export_frame = ttk.Frame(self.root)
        export_frame.pack(fill=tk.X, pady=5, padx=10)
        
        ttk.Button(export_frame, text="Exportar a CSV", 
                  command=self.export_results).pack(side=tk.LEFT, padx=5)
        ttk.Button(export_frame, text="Limpiar resultados", 
                  command=self.clear_results).pack(side=tk.LEFT, padx=5)
        
        # Barra de estado
        self.status_var = tk.StringVar(value="Listo")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def start_slave_finder(self):
        """Inicia la búsqueda de esclavos Modbus"""
        if self.slave_finding:
            return
        
        # Limpiar resultados anteriores
        self.clear_results()
        
        # Configurar cliente con timeout específico para la búsqueda
        try:
            # Crear cliente con configuración actual
            self.slave_finder_client = ModbusSerialClient(
                port=self.port_var.get(),
                baudrate=self.baudrate_var.get(),
                parity=self.parity_var.get(),
                stopbits=self.stopbits_var.get(),
                bytesize=self.bytesize_var.get(),
                timeout=self.timeout_var.get()
            )
            
            connection = self.slave_finder_client.connect()
            if not connection:
                self.update_status(f"No se pudo conectar al puerto {self.port_var.get()}")
                messagebox.showerror("Error de conexión", f"No se pudo conectar al puerto {self.port_var.get()}")
                return
            
            # Configurar variables de control
            self.slave_finding = True
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            
            # Configurar barra de progreso
            start_id = self.slave_start_var.get()
            end_id = self.slave_end_var.get()
            total_slaves = end_id - start_id + 1
            self.progress_bar["maximum"] = total_slaves
            self.progress_bar["value"] = 0
            self.progress_var.set(f"0 / {total_slaves}")
            
            # Iniciar hilo de búsqueda
            self.slave_finder_thread = threading.Thread(target=self.slave_finder_loop)
            self.slave_finder_thread.daemon = True
            self.slave_finder_thread.start()
            
            self.update_status("Búsqueda de esclavos iniciada")
            
        except Exception as e:
            self.update_status(f"Error al iniciar búsqueda: {e}")
            messagebox.showerror("Error", f"Error al iniciar búsqueda: {e}")
    
    def stop_slave_finder(self):
        """Detiene la búsqueda de esclavos Modbus"""
        if not self.slave_finding:
            return
        
        self.slave_finding = False
        if self.slave_finder_client:
            self.slave_finder_client.close()
        
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.update_status("Búsqueda de esclavos detenida")
    
    def slave_finder_loop(self):
        """Bucle para buscar esclavos Modbus"""
        start_id = self.slave_start_var.get()
        end_id = self.slave_end_var.get()
        test_function = self.slave_test_function_var.get()
        test_register = self.slave_test_register_var.get()
        progress = 0
        found_count = 0
        
        for slave_id in range(start_id, end_id + 1):
            if not self.slave_finding:
                break
            
            try:
                # Medir tiempo de respuesta
                start_time = time.time()
                
                # Leer según el tipo de registro
                if test_function == "holding":
                    response = self.slave_finder_client.read_holding_registers(test_register, 1, slave=slave_id)
                elif test_function == "input":
                    response = self.slave_finder_client.read_input_registers(test_register, 1, slave=slave_id)
                elif test_function == "coil":
                    response = self.slave_finder_client.read_coils(test_register, 1, slave=slave_id)
                elif test_function == "discrete_input":
                    response = self.slave_finder_client.read_discrete_inputs(test_register, 1, slave=slave_id)
                
                end_time = time.time()
                response_time = (end_time - start_time) * 1000  # Convertir a ms
                
                # Verificar si hay respuesta válida
                if not hasattr(response, 'isError') or not response.isError():
                    # Obtener valor del registro
                    if test_function in ["holding", "input"]:
                        value = response.registers[0] if response.registers else "N/A"
                    else:
                        value = response.bits[0] if response.bits else "N/A"
                    
                    # Añadir a la tabla en el hilo principal
                    self.root.after(0, self.add_slave_to_results, slave_id, response_time, value, "Activo")
                    found_count += 1
                else:
                    # Añadir error a la tabla
                    self.root.after(0, self.add_slave_to_results, slave_id, response_time, "N/A", f"Error: {response}")
            
            except Exception as e:
                # Ignorar errores (timeouts esperados para IDs no existentes)
                pass
            
            # Actualizar progreso
            progress += 1
            self.root.after(0, self.update_progress, progress, found_count)
            
            # Pequeña pausa para no saturar el puerto
            time.sleep(0.01)
        
        # Finalizar búsqueda
        if self.slave_finding:  # Si no fue detenido manualmente
            self.root.after(0, self.stop_slave_finder)
            self.root.after(0, self.update_status, f"Búsqueda completada. Encontrados: {found_count} esclavos")
    
    def add_slave_to_results(self, slave_id, response_time, value, status):
        """Añade un esclavo encontrado a la tabla de resultados"""
        self.results_tree.insert("", "end", values=(
            slave_id,
            f"{response_time:.1f}",
            value,
            status
        ))
    
    def update_progress(self, value, found_count):
        """Actualiza la barra de progreso de la búsqueda"""
        self.progress_bar["value"] = value
        total = self.slave_end_var.get() - self.slave_start_var.get() + 1
        self.progress_var.set(f"{value} / {total} (Encontrados: {found_count})")
    
    def clear_results(self):
        """Limpia los resultados de la búsqueda de esclavos"""
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
    
    def export_results(self):
        """Exporta los resultados de la búsqueda de esclavos a un archivo CSV"""
        if not self.results_tree.get_children():
            messagebox.showinfo("Información", "No hay datos para exportar")
            return
        
        try:
            filename = f"slave_finder_results_{time.strftime('%Y%m%d_%H%M%S')}.csv"
            filepath = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
            
            with open(filepath, 'w') as f:
                f.write("ID Esclavo,Tiempo de respuesta (ms),Valor del registro,Estado\n")
                
                for item in self.results_tree.get_children():
                    values = self.results_tree.item(item, 'values')
                    f.write(f"{values[0]},{values[1]},{values[2]},{values[3]}\n")
            
            self.update_status(f"Datos exportados a {filename}")
            messagebox.showinfo("Exportación exitosa", f"Datos exportados a {filename}")
        
        except Exception as e:
            self.update_status(f"Error al exportar datos: {e}")
            messagebox.showerror("Error", f"Error al exportar datos: {e}")
    
    def update_status(self, message):
        """Actualiza la barra de estado con un nuevo mensaje"""
        self.status_var.set(message)
    
    def on_closing(self):
        """Maneja el cierre de la aplicación"""
        # Detener hilos activos
        if self.slave_finding:
            self.stop_slave_finder()
        
        # Cerrar la aplicación
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = ModbusSlaveFinderApp(root)
    root.mainloop()