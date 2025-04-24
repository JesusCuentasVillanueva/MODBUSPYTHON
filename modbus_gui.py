import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import time
import threading
from pymodbus.client import ModbusSerialClient
from pymodbus.exceptions import ModbusException
import logging
import json
import os
from datetime import datetime

# Configurar logging
logging.basicConfig()
log = logging.getLogger()
log.setLevel(logging.INFO)

class ModbusRTUApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Modbus RTU Tool")
        self.root.geometry("900x600")
        self.root.resizable(True, True)
        
        # Variables para la configuración
        self.port_var = tk.StringVar(value="COM3")
        self.baudrate_var = tk.IntVar(value=9600)
        self.parity_var = tk.StringVar(value="N")
        self.stopbits_var = tk.IntVar(value=1)
        self.bytesize_var = tk.IntVar(value=8)
        self.timeout_var = tk.DoubleVar(value=1.0)
        self.slave_var = tk.IntVar(value=0)
        
        # Variables para operaciones
        self.register_var = tk.IntVar(value=0)
        self.register_type_var = tk.StringVar(value="holding")
        self.scale_var = tk.DoubleVar(value=0.1)
        self.count_var = tk.IntVar(value=1)
        self.write_value_var = tk.StringVar(value="0")
        self.auto_refresh_var = tk.BooleanVar(value=False)
        self.refresh_rate_var = tk.DoubleVar(value=1.0)
        
        # Variables para monitoreo
        self.monitoring = False
        self.monitor_thread = None
        self.client = None
        self.read_values = []
        
        # Historial de comandos
        self.command_history = []
        self.history_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "modbus_history.json")
        self.load_history()
        
        # Crear interfaz
        self.create_widgets()
        
        # Configurar cierre de la aplicación
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def create_widgets(self):
        # Frame principal con pestañas
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Pestaña 1: Operaciones básicas
        operations_frame = ttk.Frame(notebook, padding=10)
        notebook.add(operations_frame, text="Operaciones")
        
        # Pestaña 2: Monitor continuo
        monitor_frame = ttk.Frame(notebook, padding=10)
        notebook.add(monitor_frame, text="Monitor")
        
        # Pestaña 3: Escáner de registros
        scanner_frame = ttk.Frame(notebook, padding=10)
        notebook.add(scanner_frame, text="Escáner")
        
        # Pestaña 4: Historial
        history_frame = ttk.Frame(notebook, padding=10)
        notebook.add(history_frame, text="Historial")
        
        # Configurar cada pestaña
        self.setup_operations_tab(operations_frame)
        self.setup_monitor_tab(monitor_frame)
        self.setup_scanner_tab(scanner_frame)
        self.setup_history_tab(history_frame)
        
        # Barra de estado
        self.status_var = tk.StringVar(value="Listo")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def setup_operations_tab(self, parent):
        # Dividir en dos paneles
        left_frame = ttk.Frame(parent)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 5))
        
        right_frame = ttk.Frame(parent)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Panel izquierdo: Configuración
        config_frame = ttk.LabelFrame(left_frame, text="Configuración de comunicación", padding=10)
        config_frame.pack(fill=tk.X, expand=False, pady=(0, 10))
        
        ttk.Label(config_frame, text="Puerto:").grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Entry(config_frame, textvariable=self.port_var, width=15).grid(row=0, column=1, sticky=tk.W, pady=2)
        
        ttk.Label(config_frame, text="Baudrate:").grid(row=1, column=0, sticky=tk.W, pady=2)
        ttk.Combobox(config_frame, textvariable=self.baudrate_var, 
                    values=[1200, 2400, 4800, 9600, 19200, 38400, 57600, 115200], 
                    width=13).grid(row=1, column=1, sticky=tk.W, pady=2)
        
        ttk.Label(config_frame, text="Paridad:").grid(row=2, column=0, sticky=tk.W, pady=2)
        ttk.Combobox(config_frame, textvariable=self.parity_var, 
                    values=["N", "E", "O"], width=13).grid(row=2, column=1, sticky=tk.W, pady=2)
        
        ttk.Label(config_frame, text="Bits de parada:").grid(row=3, column=0, sticky=tk.W, pady=2)
        ttk.Combobox(config_frame, textvariable=self.stopbits_var, 
                    values=[1, 2], width=13).grid(row=3, column=1, sticky=tk.W, pady=2)
        
        ttk.Label(config_frame, text="Bits de datos:").grid(row=4, column=0, sticky=tk.W, pady=2)
        ttk.Combobox(config_frame, textvariable=self.bytesize_var, 
                    values=[7, 8], width=13).grid(row=4, column=1, sticky=tk.W, pady=2)
        
        ttk.Label(config_frame, text="Timeout (s):").grid(row=5, column=0, sticky=tk.W, pady=2)
        ttk.Entry(config_frame, textvariable=self.timeout_var, width=15).grid(row=5, column=1, sticky=tk.W, pady=2)
        
        ttk.Label(config_frame, text="Dirección esclavo:").grid(row=6, column=0, sticky=tk.W, pady=2)
        ttk.Entry(config_frame, textvariable=self.slave_var, width=15).grid(row=6, column=1, sticky=tk.W, pady=2)
        
        # Panel izquierdo: Operaciones
        operations_frame = ttk.LabelFrame(left_frame, text="Operaciones Modbus", padding=10)
        operations_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(operations_frame, text="Tipo de registro:").grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Combobox(operations_frame, textvariable=self.register_type_var, 
                    values=["holding", "input", "coil", "discrete_input"], 
                    width=13).grid(row=0, column=1, sticky=tk.W, pady=2)
        
        ttk.Label(operations_frame, text="Dirección registro:").grid(row=1, column=0, sticky=tk.W, pady=2)
        ttk.Entry(operations_frame, textvariable=self.register_var, width=15).grid(row=1, column=1, sticky=tk.W, pady=2)
        
        ttk.Label(operations_frame, text="Cantidad:").grid(row=2, column=0, sticky=tk.W, pady=2)
        ttk.Entry(operations_frame, textvariable=self.count_var, width=15).grid(row=2, column=1, sticky=tk.W, pady=2)
        
        ttk.Label(operations_frame, text="Factor de escala:").grid(row=3, column=0, sticky=tk.W, pady=2)
        ttk.Combobox(operations_frame, textvariable=self.scale_var, 
                    values=[1.0, 0.1, 0.01, 0.001], width=13).grid(row=3, column=1, sticky=tk.W, pady=2)
        
        ttk.Separator(operations_frame, orient=tk.HORIZONTAL).grid(row=4, column=0, columnspan=2, sticky=tk.EW, pady=10)
        
        ttk.Label(operations_frame, text="Valor a escribir:").grid(row=5, column=0, sticky=tk.W, pady=2)
        ttk.Entry(operations_frame, textvariable=self.write_value_var, width=15).grid(row=5, column=1, sticky=tk.W, pady=2)
        
        # Botones de operaciones
        button_frame = ttk.Frame(operations_frame)
        button_frame.grid(row=6, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Leer", command=self.read_registers).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Escribir", command=self.write_register).pack(side=tk.LEFT, padx=5)
        
        # Panel derecho: Resultados
        results_frame = ttk.LabelFrame(right_frame, text="Resultados", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True)
        
        self.results_text = scrolledtext.ScrolledText(results_frame, wrap=tk.WORD, width=50, height=20)
        self.results_text.pack(fill=tk.BOTH, expand=True)
        self.results_text.config(state=tk.DISABLED)
        
        # Botones de control para resultados
        control_frame = ttk.Frame(results_frame)
        control_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(control_frame, text="Limpiar", command=self.clear_results).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(control_frame, text="Auto-refresh", variable=self.auto_refresh_var, 
                       command=self.toggle_auto_refresh).pack(side=tk.LEFT, padx=5)
        ttk.Label(control_frame, text="Intervalo (s):").pack(side=tk.LEFT, padx=5)
        ttk.Entry(control_frame, textvariable=self.refresh_rate_var, width=5).pack(side=tk.LEFT, padx=5)
    
    def setup_monitor_tab(self, parent):
        # Panel superior: Configuración del monitor
        config_frame = ttk.LabelFrame(parent, text="Configuración del monitor", padding=10)
        config_frame.pack(fill=tk.X, expand=False, pady=(0, 10))
        
        # Fila 1
        frame1 = ttk.Frame(config_frame)
        frame1.pack(fill=tk.X, pady=5)
        
        ttk.Label(frame1, text="Registro:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(frame1, textvariable=self.register_var, width=8).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(frame1, text="Tipo:").pack(side=tk.LEFT, padx=5)
        ttk.Combobox(frame1, textvariable=self.register_type_var, 
                    values=["holding", "input"], width=10).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(frame1, text="Escala:").pack(side=tk.LEFT, padx=5)
        ttk.Combobox(frame1, textvariable=self.scale_var, 
                    values=[1.0, 0.1, 0.01, 0.001], width=8).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(frame1, text="Intervalo (s):").pack(side=tk.LEFT, padx=5)
        ttk.Entry(frame1, textvariable=self.refresh_rate_var, width=5).pack(side=tk.LEFT, padx=5)
        
        # Botones de control
        self.monitor_start_button = ttk.Button(frame1, text="Iniciar Monitor", command=self.start_monitoring)
        self.monitor_start_button.pack(side=tk.LEFT, padx=5)
        
        self.monitor_stop_button = ttk.Button(frame1, text="Detener", command=self.stop_monitoring, state=tk.DISABLED)
        self.monitor_stop_button.pack(side=tk.LEFT, padx=5)
        
        # Panel inferior: Visualización
        display_frame = ttk.LabelFrame(parent, text="Monitor de valores", padding=10)
        display_frame.pack(fill=tk.BOTH, expand=True)
        
        # Valor actual grande
        self.current_value_var = tk.StringVar(value="--")
        self.current_value_label = ttk.Label(display_frame, textvariable=self.current_value_var, font=("Arial", 36))
        self.current_value_label.pack(pady=10)
        
        # Gráfico/tabla de valores históricos
        history_frame = ttk.Frame(display_frame)
        history_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Tabla para el historial
        columns = ("timestamp", "value", "formatted")
        self.monitor_tree = ttk.Treeview(history_frame, columns=columns, show="headings")
        self.monitor_tree.heading("timestamp", text="Hora")
        self.monitor_tree.heading("value", text="Valor (raw)")
        self.monitor_tree.heading("formatted", text="Valor formateado")
        self.monitor_tree.column("timestamp", width=100)
        self.monitor_tree.column("value", width=100)
        self.monitor_tree.column("formatted", width=150)
        self.monitor_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Scrollbar para la tabla
        scrollbar = ttk.Scrollbar(history_frame, orient=tk.VERTICAL, command=self.monitor_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.monitor_tree.configure(yscrollcommand=scrollbar.set)
        
        # Botones para exportar datos
        export_frame = ttk.Frame(display_frame)
        export_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(export_frame, text="Exportar a CSV", command=self.export_monitor_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(export_frame, text="Limpiar historial", command=self.clear_monitor_history).pack(side=tk.LEFT, padx=5)
    
    def setup_scanner_tab(self, parent):
        # Panel superior: Configuración del escáner
        config_frame = ttk.LabelFrame(parent, text="Configuración del escáner", padding=10)
        config_frame.pack(fill=tk.X, expand=False, pady=(0, 10))
        
        # Fila 1
        frame1 = ttk.Frame(config_frame)
        frame1.pack(fill=tk.X, pady=5)
        
        ttk.Label(frame1, text="Registro inicial:").pack(side=tk.LEFT, padx=5)
        self.scan_start_var = tk.IntVar(value=0)
        ttk.Entry(frame1, textvariable=self.scan_start_var, width=8).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(frame1, text="Cantidad:").pack(side=tk.LEFT, padx=5)
        self.scan_count_var = tk.IntVar(value=100)
        ttk.Entry(frame1, textvariable=self.scan_count_var, width=8).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(frame1, text="Tipo:").pack(side=tk.LEFT, padx=5)
        self.scan_type_var = tk.StringVar(value="holding")
        ttk.Combobox(frame1, textvariable=self.scan_type_var, 
                    values=["holding", "input", "both"], width=10).pack(side=tk.LEFT, padx=5)
        
        # Botones de control
        ttk.Button(frame1, text="Iniciar Escaneo", command=self.start_scanning).pack(side=tk.LEFT, padx=5)
        
        # Panel inferior: Resultados del escaneo
        results_frame = ttk.LabelFrame(parent, text="Resultados del escaneo", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True)
        
        # Tabla para los resultados
        columns = ("address", "type", "value_dec", "value_hex", "value_scaled")
        self.scan_tree = ttk.Treeview(results_frame, columns=columns, show="headings")
        self.scan_tree.heading("address", text="Dirección")
        self.scan_tree.heading("type", text="Tipo")
        self.scan_tree.heading("value_dec", text="Valor (Dec)")
        self.scan_tree.heading("value_hex", text="Valor (Hex)")
        self.scan_tree.heading("value_scaled", text="Valor (÷10)")
        self.scan_tree.column("address", width=80)
        self.scan_tree.column("type", width=80)
        self.scan_tree.column("value_dec", width=100)
        self.scan_tree.column("value_hex", width=100)
        self.scan_tree.column("value_scaled", width=100)
        self.scan_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Scrollbar para la tabla
        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.scan_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.scan_tree.configure(yscrollcommand=scrollbar.set)
        
        # Botones para exportar datos
        export_frame = ttk.Frame(results_frame)
        export_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(export_frame, text="Exportar a CSV", command=self.export_scan_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(export_frame, text="Limpiar resultados", command=self.clear_scan_results).pack(side=tk.LEFT, padx=5)
    
    def setup_history_tab(self, parent):
        # Panel para el historial de comandos
        history_frame = ttk.LabelFrame(parent, text="Historial de comandos", padding=10)
        history_frame.pack(fill=tk.BOTH, expand=True)
        
        # Tabla para el historial
        columns = ("timestamp", "operation", "parameters", "result")
        self.history_tree = ttk.Treeview(history_frame, columns=columns, show="headings")
        self.history_tree.heading("timestamp", text="Fecha/Hora")
        self.history_tree.heading("operation", text="Operación")
        self.history_tree.heading("parameters", text="Parámetros")
        self.history_tree.heading("result", text="Resultado")
        self.history_tree.column("timestamp", width=150)
        self.history_tree.column("operation", width=100)
        self.history_tree.column("parameters", width=300)
        self.history_tree.column("result", width=200)
        self.history_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Scrollbar para la tabla
        scrollbar = ttk.Scrollbar(history_frame, orient=tk.VERTICAL, command=self.history_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.history_tree.configure(yscrollcommand=scrollbar.set)
        
        # Botones para el historial
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(button_frame, text="Repetir comando seleccionado", 
                  command=self.repeat_command).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Limpiar historial", 
                  command=self.clear_history).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Exportar historial", 
                  command=self.export_history).pack(side=tk.LEFT, padx=5)
    
    def create_modbus_client(self):
        """Crea y conecta un cliente Modbus RTU con la configuración actual"""
        try:
            client = ModbusSerialClient(
                port=self.port_var.get(),
                baudrate=self.baudrate_var.get(),
                parity=self.parity_var.get(),
                stopbits=self.stopbits_var.get(),
                bytesize=self.bytesize_var.get(),
                timeout=self.timeout_var.get()
            )
            
            connection = client.connect()
            if not connection:
                self.update_status(f"No se pudo conectar al puerto {self.port_var.get()}")
                return None
            
            self.update_status(f"Conectado a {self.port_var.get()}")
            return client
        
        except Exception as e:
            self.update_status(f"Error al crear cliente Modbus: {e}")
            return None
    
    def read_registers(self):
        """Lee registros Modbus según la configuración actual"""
        client = self.create_modbus_client()
        if not client:
            return
        
        try:
            register = self.register_var.get()
            count = self.count_var.get()
            slave = self.slave_var.get()
            register_type = self.register_type_var.get()
            scale = self.scale_var.get()
            
            # Leer según el tipo de registro
            if register_type == "holding":
                response = client.read_holding_registers(register, count, slave=slave)
            elif register_type == "input":
                response = client.read_input_registers(register, count, slave=slave)
            elif register_type == "coil":
                response = client.read_coils(register, count, slave=slave)
            elif register_type == "discrete_input":
                response = client.read_discrete_inputs(register, count, slave=slave)
            else:
                self.update_results(f"Tipo de registro no válido: {register_type}")
                client.close()
                return
            
            # Procesar respuesta
            if not hasattr(response, 'isError') or not response.isError():
                if register_type in ["holding", "input"]:
                    self.update_results(f"Lectura exitosa de {count} registros {register_type} desde {register}:")
                    for i, value in enumerate(response.registers):
                        addr = register + i
                        scaled_value = value * scale
                        
                        # Formatear según escala
                        if scale == 1:
                            formatted = f"{scaled_value}"
                        elif scale == 0.1:
                            formatted = f"{scaled_value:.1f}"
                        elif scale == 0.01:
                            formatted = f"{scaled_value:.2f}"
                        else:
                            formatted = f"{scaled_value:.3f}"
                        
                        self.update_results(f"  Registro {addr}: {value} (0x{value:04X}) → {formatted}")
                else:
                    self.update_results(f"Lectura exitosa de {count} bits {register_type} desde {register}:")
                    for i, value in enumerate(response.bits):
                        addr = register + i
                        self.update_results(f"  Bit {addr}: {value}")
                
                # Guardar en historial
                self.add_to_history("Lectura", 
                                   f"Tipo: {register_type}, Reg: {register}, Count: {count}, Slave: {slave}",
                                   "Exitoso")
            else:
                self.update_results(f"Error al leer registros: {response}")
                self.add_to_history("Lectura", 
                                   f"Tipo: {register_type}, Reg: {register}, Count: {count}, Slave: {slave}",
                                   f"Error: {response}")
        
        except Exception as e:
            self.update_results(f"Error: {e}")
            self.add_to_history("Lectura", 
                               f"Tipo: {register_type}, Reg: {register}, Count: {count}, Slave: {slave}",
                               f"Excepción: {e}")
        
        finally:
            client.close()
    
    def write_register(self):
        """Escribe en un registro Modbus según la configuración actual"""
        client = self.create_modbus_client()
        if not client:
            return
        
        try:
            register = self.register_var.get()
            slave = self.slave_var.get()
            register_type = self.register_type_var.get()
            value_str = self.write_value_var.get()
            
            # Convertir valor según el tipo de registro
            if register_type in ["holding", "input"]:
                try:
                    # Verificar si es un valor decimal con punto
                    if '.' in value_str:
                        float_value = float(value_str)
                        scale = self.scale_var.get()
                        value = int(float_value / scale)
                    else:
                        value = int(value_str)
                except ValueError:
                    self.update_results(f"Valor no válido: {value_str}")
                    client.close()
                    return
            else:
                # Para coils, convertir a booleano
                value = value_str.lower() in ['1', 'true', 't', 'yes', 'y', 'on']
            
            # Escribir según el tipo de registro
            if register_type == "holding":
                response = client.write_register(register, value, slave=slave)
            elif register_type == "coil":
                response = client.write_coil(register, value, slave=slave)
            else:
                self.update_results(f"No se puede escribir en registro tipo: {register_type}")
                client.close()
                return
            
            # Procesar respuesta
            if not hasattr(response, 'isError') or not response.isError():
                self.update_results(f"Escritura exitosa en registro {register_type} {register}: {value}")
                self.add_to_history("Escritura", 
                                   f"Tipo: {register_type}, Reg: {register}, Valor: {value}, Slave: {slave}",
                                   "Exitoso")
            else:
                self.update_results(f"Error al escribir en registro: {response}")
                self.add_to_history("Escritura", 
                                   f"Tipo: {register_type}, Reg: {register}, Valor: {value}, Slave: {slave}",
                                   f"Error: {response}")
        
        except Exception as e:
            self.update_results(f"Error: {e}")
            self.add_to_history("Escritura", 
                               f"Tipo: {register_type}, Reg: {register}, Valor: {value_str}, Slave: {slave}",
                               f"Excepción: {e}")
        
        finally:
            client.close()
    
    def toggle_auto_refresh(self):
        """Activa o desactiva la actualización automática de lecturas"""
        if self.auto_refresh_var.get():
            self.start_auto_refresh()
        else:
            self.stop_auto_refresh()
    
    def start_auto_refresh(self):
        """Inicia la actualización automática de lecturas"""
        if hasattr(self, 'refresh_thread') and self.refresh_thread and self.refresh_thread.is_alive():
            return
        
        self.auto_refreshing = True
        self.refresh_thread = threading.Thread(target=self.auto_refresh_loop)
        self.refresh_thread.daemon = True
        self.refresh_thread.start()
        self.update_status("Auto-refresh activado")
    
    def stop_auto_refresh(self):
        """Detiene la actualización automática de lecturas"""
        self.auto_refreshing = False
        self.update_status("Auto-refresh desactivado")
    
    def auto_refresh_loop(self):
        """Bucle para actualización automática de lecturas"""
        while self.auto_refreshing:
            self.read_registers()
            time.sleep(self.refresh_rate_var.get())
    
    def start_monitoring(self):
        """Inicia el monitoreo continuo de un registro"""
        if self.monitoring:
            return
        
        self.client = self.create_modbus_client()
        if not self.client:
            return
        
        self.monitoring = True
        self.monitor_start_button.config(state=tk.DISABLED)
        self.monitor_stop_button.config(state=tk.NORMAL)
        
        # Iniciar hilo de monitoreo
        self.monitor_thread = threading.Thread(target=self.monitor_loop)
        self.monitor_thread.daemon = True
        self.monitor_thread.start()
        
        self.update_status("Monitoreo iniciado")
    
    def stop_monitoring(self):
        """Detiene el monitoreo continuo"""
        if not self.monitoring:
            return
        
        self.monitoring = False
        if self.client:
            self.client.close()
        
        self.monitor_start_button.config(state=tk.NORMAL)
        self.monitor_stop_button.config(state=tk.DISABLED)
        self.update_status("Monitoreo detenido")
    
    def monitor_loop(self):
        """Bucle para monitoreo continuo de un registro"""
        register = self.register_var.get()
        register_type = self.register_type_var.get()
        scale = self.scale_var.get()
        slave = self.slave_var.get()
        max_history = 100  # Máximo número de entradas en el historial
        
        while self.monitoring:
            try:
                # Leer registro según el tipo
                if register_type == "holding":
                    response = self.client.read_holding_registers(register, 1, slave=slave)
                elif register_type == "input":
                    response = self.client.read_input_registers(register, 1, slave=slave)
                else:
                    self.root.after(0, self.update_status, f"Tipo de registro no válido: {register_type}")
                    break
                
                if not hasattr(response, 'isError') or not response.isError():
                    value = response.registers[0]
                    scaled_value = value * scale
                    
                    # Formatear según escala
                    if scale == 1:
                        formatted = f"{scaled_value}"
                    elif scale == 0.1:
                        formatted = f"{scaled_value:.1f}"
                    elif scale == 0.01:
                        formatted = f"{scaled_value:.2f}"
                    else:
                        formatted = f"{scaled_value:.3f}"
                    
                    # Actualizar interfaz en el hilo principal
                    self.root.after(0, self.update_monitor_display, value, scaled_value, formatted)
                else:
                    self.root.after(0, self.update_status, f"Error al leer registro: {response}")
            
            except Exception as e:
                self.root.after(0, self.update_status, f"Error: {e}")
                if not self.monitoring:  # Si se detuvo durante la excepción
                    break
            
            time.sleep(self.refresh_rate_var.get())
    
    def update_monitor_display(self, value, scaled_value, formatted):
        """Actualiza la visualización del monitor con un nuevo valor"""
        timestamp = time.strftime("%H:%M:%S")
        
        # Actualizar etiqueta de valor actual
        if formatted.endswith(".0"):
            display_value = formatted.rstrip("0").rstrip(".")
        else:
            display_value = formatted
        self.current_value_var.set(f"{display_value}")
        
        # Añadir a la tabla de historial
        self.monitor_tree.insert("", 0, values=(timestamp, value, formatted))
        
        # Limitar el número de entradas en la tabla
        if len(self.monitor_tree.get_children()) > 100:
            last_item = self.monitor_tree.get_children()[-1]
            self.monitor_tree.delete(last_item)
    
    def start_scanning(self):
        """Inicia el escaneo de un rango de registros"""
        client = self.create_modbus_client()
        if not client:
            return
        
        # Limpiar resultados anteriores
        self.clear_scan_results()
        
        try:
            start_register = self.scan_start_var.get()
            count = self.scan_count_var.get()
            scan_type = self.scan_type_var.get()
            slave = self.slave_var.get()
            
            self.update_status(f"Escaneando {count} registros desde {start_register}...")
            
            # Limitar el número máximo de registros por seguridad
            if count > 125:
                if not messagebox.askyesno("Advertencia", 
                                         f"Estás intentando escanear {count} registros, lo que puede llevar tiempo. ¿Continuar?"):
                    client.close()
                    self.update_status("Escaneo cancelado")
                    return
            
            # Escanear registros holding
            if scan_type in ["holding", "both"]:
                for i in range(0, count, 20):  # Leer en bloques de 20 para evitar timeouts
                    block_size = min(20, count - i)
                    try:
                        response = client.read_holding_registers(start_register + i, block_size, slave=slave)
                        if not hasattr(response, 'isError') or not response.isError():
                            for j, value in enumerate(response.registers):
                                addr = start_register + i + j
                                self.scan_tree.insert("", "end", values=(
                                    addr, 
                                    "holding", 
                                    value, 
                                    f"0x{value:04X}", 
                                    f"{value/10.0:.1f}"
                                ))
                    except Exception as e:
                        pass  # Ignorar errores en escaneo
                    
                    # Actualizar la interfaz
                    self.root.update()
            
            # Escanear registros input
            if scan_type in ["input", "both"]:
                for i in range(0, count, 20):
                    block_size = min(20, count - i)
                    try:
                        response = client.read_input_registers(start_register + i, block_size, slave=slave)
                        if not hasattr(response, 'isError') or not response.isError():
                            for j, value in enumerate(response.registers):
                                addr = start_register + i + j
                                self.scan_tree.insert("", "end", values=(
                                    addr, 
                                    "input", 
                                    value, 
                                    f"0x{value:04X}", 
                                    f"{value/10.0:.1f}"
                                ))
                    except Exception as e:
                        pass  # Ignorar errores en escaneo
                    
                    # Actualizar la interfaz
                    self.root.update()
            
            self.update_status(f"Escaneo completado")
            
        except Exception as e:
            self.update_status(f"Error durante el escaneo: {e}")
        
        finally:
            client.close()
    
    def clear_scan_results(self):
        """Limpia los resultados del escaneo"""
        for item in self.scan_tree.get_children():
            self.scan_tree.delete(item)
    
    def export_scan_data(self):
        """Exporta los resultados del escaneo a un archivo CSV"""
        if not self.scan_tree.get_children():
            messagebox.showinfo("Información", "No hay datos para exportar")
            return
        
        try:
            filename = f"scan_results_{time.strftime('%Y%m%d_%H%M%S')}.csv"
            filepath = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
            
            with open(filepath, 'w') as f:
                f.write("Dirección,Tipo,Valor (Dec),Valor (Hex),Valor (÷10)\n")
                
                for item in self.scan_tree.get_children():
                    values = self.scan_tree.item(item, 'values')
                    f.write(f"{values[0]},{values[1]},{values[2]},{values[3]},{values[4]}\n")
            
            self.update_status(f"Datos exportados a {filename}")
            messagebox.showinfo("Exportación exitosa", f"Datos exportados a {filename}")
        
        except Exception as e:
            self.update_status(f"Error al exportar datos: {e}")
            messagebox.showerror("Error", f"Error al exportar datos: {e}")
    
    def export_monitor_data(self):
        """Exporta los datos del monitor a un archivo CSV"""
        if not self.monitor_tree.get_children():
            messagebox.showinfo("Información", "No hay datos para exportar")
            return
        
        try:
            filename = f"monitor_data_{time.strftime('%Y%m%d_%H%M%S')}.csv"
            filepath = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
            
            with open(filepath, 'w') as f:
                f.write("Hora,Valor (raw),Valor formateado\n")
                
                for item in self.monitor_tree.get_children():
                    values = self.monitor_tree.item(item, 'values')
                    f.write(f"{values[0]},{values[1]},{values[2]}\n")
            
            self.update_status(f"Datos exportados a {filename}")
            messagebox.showinfo("Exportación exitosa", f"Datos exportados a {filename}")
        
        except Exception as e:
            self.update_status(f"Error al exportar datos: {e}")
            messagebox.showerror("Error", f"Error al exportar datos: {e}")
    
    def clear_monitor_history(self):
        """Limpia el historial del monitor"""
        for item in self.monitor_tree.get_children():
            self.monitor_tree.delete(item)
        self.current_value_var.set("--")
    
    def update_results(self, message):
        """Actualiza el área de resultados con un nuevo mensaje"""
        self.results_text.config(state=tk.NORMAL)
        self.results_text.insert(tk.END, message + "\n")
        self.results_text.see(tk.END)
        self.results_text.config(state=tk.DISABLED)
    
    def clear_results(self):
        """Limpia el área de resultados"""
        self.results_text.config(state=tk.NORMAL)
        self.results_text.delete(1.0, tk.END)
        self.results_text.config(state=tk.DISABLED)
    
    def update_status(self, message):
        """Actualiza la barra de estado con un nuevo mensaje"""
        self.status_var.set(message)
    
    def add_to_history(self, operation, parameters, result):
        """Añade una entrada al historial de comandos"""
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        
        # Añadir a la lista interna
        self.command_history.append({
            "timestamp": timestamp,
            "operation": operation,
            "parameters": parameters,
            "result": result
        })
        
        # Añadir a la tabla de historial
        self.history_tree.insert("", 0, values=(timestamp, operation, parameters, result))
        
        # Guardar historial en archivo
        self.save_history()
    
    def load_history(self):
        """Carga el historial de comandos desde un archivo"""
        try:
            if os.path.exists(self.history_file):
                with open(self.history_file, 'r') as f:
                    self.command_history = json.load(f)
                
                # Cargar en la tabla (si ya está creada)
                if hasattr(self, 'history_tree'):
                    for item in self.command_history:
                        self.history_tree.insert("", "end", values=(
                            item["timestamp"],
                            item["operation"],
                            item["parameters"],
                            item["result"]
                        ))
        except Exception as e:
            print(f"Error al cargar historial: {e}")
    
    def save_history(self):
        """Guarda el historial de comandos en un archivo"""
        try:
            with open(self.history_file, 'w') as f:
                json.dump(self.command_history, f)
        except Exception as e:
            print(f"Error al guardar historial: {e}")
    
    def clear_history(self):
        """Limpia el historial de comandos"""
        if messagebox.askyesno("Confirmar", "¿Estás seguro de que quieres borrar todo el historial?"):
            self.command_history = []
            for item in self.history_tree.get_children():
                self.history_tree.delete(item)
            self.save_history()
    
    def export_history(self):
        """Exporta el historial de comandos a un archivo CSV"""
        if not self.command_history:
            messagebox.showinfo("Información", "No hay historial para exportar")
            return
        
        try:
            filename = f"command_history_{time.strftime('%Y%m%d_%H%M%S')}.csv"
            filepath = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
            
            with open(filepath, 'w') as f:
                f.write("Fecha/Hora,Operación,Parámetros,Resultado\n")
                
                for item in self.command_history:
                    f.write(f"{item['timestamp']},{item['operation']},{item['parameters']},{item['result']}\n")
            
            self.update_status(f"Historial exportado a {filename}")
            messagebox.showinfo("Exportación exitosa", f"Historial exportado a {filename}")
        
        except Exception as e:
            self.update_status(f"Error al exportar historial: {e}")
            messagebox.showerror("Error", f"Error al exportar historial: {e}")
    
    def repeat_command(self):
        """Repite el comando seleccionado en el historial"""
        selected = self.history_tree.selection()
        if not selected:
            messagebox.showinfo("Información", "Selecciona un comando del historial para repetir")
            return
        
        values = self.history_tree.item(selected[0], 'values')
        operation = values[1]
        parameters = values[2]
        
        # Extraer parámetros
        params = {}
        for param in parameters.split(', '):
            if ': ' in param:
                key, value = param.split(': ', 1)
                params[key] = value
        
        # Configurar la interfaz según los parámetros
        if 'Tipo' in params:
            self.register_type_var.set(params['Tipo'])
        if 'Reg' in params:
            self.register_var.set(int(params['Reg']))
        if 'Count' in params:
            self.count_var.set(int(params['Count']))
        if 'Slave' in params:
            self.slave_var.set(int(params['Slave']))
        if 'Valor' in params:
            self.write_value_var.set(params['Valor'])
        
        # Ejecutar la operación
        if operation == "Lectura":
            self.read_registers()
        elif operation == "Escritura":
            self.write_register()
    
    def on_closing(self):
        """Maneja el cierre de la aplicación"""
        # Detener hilos activos
        if self.monitoring:
            self.stop_monitoring()
        
        if hasattr(self, 'auto_refreshing') and self.auto_refreshing:
            self.stop_auto_refresh()
        
        # Guardar historial
        self.save_history()
        
        # Cerrar la aplicación
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = ModbusRTUApp(root)
    root.mainloop()
