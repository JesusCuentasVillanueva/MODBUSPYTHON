import time
from pymodbus.client import ModbusSerialClient
from pymodbus.exceptions import ModbusException
import logging

# Configurar logging para ver detalles de la comunicación
logging.basicConfig()
log = logging.getLogger()
log.setLevel(logging.INFO)

def scan_modbus_registers(port='COM3', baudrate=9600, parity='N', stopbits=1, bytesize=8, timeout=1, slave_address=0):
    """
    Escanea registros Modbus para encontrar valores de temperatura del sensor PT1000.
    
    Args:
        port (str): Puerto COM del convertidor RS485
        baudrate (int): Velocidad de comunicación
        parity (str): Paridad ('N' para ninguna, 'E' para par, 'O' para impar)
        stopbits (int): Bits de parada
        bytesize (int): Tamaño de byte
        timeout (int): Tiempo de espera para la respuesta
        slave_address (int): Dirección del esclavo Modbus
    """
    print(f"Iniciando escaneo de registros Modbus RTU en {port}")
    print(f"Configuración: {baudrate} baudios, {bytesize}{parity}{stopbits}, timeout {timeout}s")
    print(f"Dirección del esclavo: {slave_address}")
    
    # Crear cliente Modbus RTU
    client = ModbusSerialClient(
        port=port,
        baudrate=baudrate,
        parity=parity,
        stopbits=stopbits,
        bytesize=bytesize,
        timeout=timeout
    )
    
    # Intentar conectar
    connection = client.connect()
    if connection:
        print(f"Conexión establecida con el puerto {port}")
        
        try:
            # Definir rangos de registros a escanear
            register_ranges = [
                (0, 50),       # Primeros 50 registros
                (100, 50),     # Registros 100-149
                (200, 50),     # Registros 200-249
                (300, 50),     # Registros 300-349
                (400, 50),     # Registros 400-449
                (500, 50),     # Registros 500-549
                (1000, 50),    # Registros 1000-1049
                (2000, 50),    # Registros 2000-2049
                (3000, 50),    # Registros 3000-3049
            ]
            
            temperature_candidates = []
            
            # Escanear registros de retención (holding registers)
            print("\n--- Escaneando registros de retención (holding registers) ---")
            for start_register, count in register_ranges:
                try:
                    print(f"Leyendo registros {start_register}-{start_register+count-1}...")
                    response = client.read_holding_registers(start_register, count, slave=slave_address)
                    
                    if not hasattr(response, 'isError') or not response.isError():
                        print(f"Registros leídos: {response.registers}")
                        
                        # Buscar posibles valores de temperatura
                        for i, value in enumerate(response.registers):
                            reg_address = start_register + i
                            
                            # Probar diferentes escalas comunes para temperaturas
                            # Valor directo (sin escala)
                            if -50 <= value <= 150:
                                temperature_candidates.append((reg_address, value, "holding", "1x"))
                            
                            # Valor dividido por 10 (un decimal)
                            temp_val = value / 10.0
                            if -50 <= temp_val <= 150:
                                temperature_candidates.append((reg_address, temp_val, "holding", "0.1x"))
                            
                            # Valor dividido por 100 (dos decimales)
                            temp_val = value / 100.0
                            if -50 <= temp_val <= 150:
                                temperature_candidates.append((reg_address, temp_val, "holding", "0.01x"))
                except Exception as e:
                    print(f"Error al leer registros {start_register}-{start_register+count-1}: {e}")
                
                time.sleep(0.2)  # Pequeña pausa entre lecturas
            
            # Escanear registros de entrada (input registers)
            print("\n--- Escaneando registros de entrada (input registers) ---")
            for start_register, count in register_ranges:
                try:
                    print(f"Leyendo registros de entrada {start_register}-{start_register+count-1}...")
                    response = client.read_input_registers(start_register, count, slave=slave_address)
                    
                    if not hasattr(response, 'isError') or not response.isError():
                        print(f"Registros de entrada leídos: {response.registers}")
                        
                        # Buscar posibles valores de temperatura
                        for i, value in enumerate(response.registers):
                            reg_address = start_register + i
                            
                            # Probar diferentes escalas comunes para temperaturas
                            # Valor directo (sin escala)
                            if -50 <= value <= 150:
                                temperature_candidates.append((reg_address, value, "input", "1x"))
                            
                            # Valor dividido por 10 (un decimal)
                            temp_val = value / 10.0
                            if -50 <= temp_val <= 150:
                                temperature_candidates.append((reg_address, temp_val, "input", "0.1x"))
                            
                            # Valor dividido por 100 (dos decimales)
                            temp_val = value / 100.0
                            if -50 <= temp_val <= 150:
                                temperature_candidates.append((reg_address, temp_val, "input", "0.01x"))
                except Exception as e:
                    print(f"Error al leer registros de entrada {start_register}-{start_register+count-1}: {e}")
                
                time.sleep(0.2)  # Pequeña pausa entre lecturas
            
            # Mostrar resultados
            if temperature_candidates:
                print("\n=== POSIBLES VALORES DE TEMPERATURA ENCONTRADOS ===")
                print("Dirección | Valor | Tipo | Escala")
                print("-" * 50)
                for reg_address, value, reg_type, scale in temperature_candidates:
                    print(f"{reg_address:8} | {value:5.2f}°C | {reg_type:7} | {scale}")
                
                print("\nPara leer continuamente un registro específico, use el siguiente código:")
                print("python monitor_temperature.py --port COM3 --baudrate 9600 --parity N --address 0 --register X --type Y")
                print("Donde X es la dirección del registro y Y es 'holding' o 'input'")
            else:
                print("\nNo se encontraron valores que parezcan temperaturas en los rangos escaneados.")
                print("Sugerencias:")
                print("1. Verifica que el sensor PT1000 esté correctamente conectado al Danfoss MCX06D")
                print("2. Consulta la documentación del Danfoss MCX06D para conocer los registros específicos del PT1000")
                print("3. Prueba con diferentes rangos de registros")
        
        except ModbusException as e:
            print(f"\nError de comunicación Modbus: {e}")
        
        except Exception as e:
            print(f"\nError general: {e}")
        
        finally:
            client.close()
            print("\nConexión cerrada")
    else:
        print(f"\nNo se pudo establecer conexión con el puerto {port}")
        print("Verifica que:")
        print("1. El puerto COM3 esté disponible")
        print("2. El convertidor USB-RS485 esté correctamente instalado")
        print("3. No haya otro programa utilizando el puerto")

def test_modbus_connection(port='COM3', baudrate=9600, parity='N', stopbits=1, bytesize=8, timeout=1):
    """
    Prueba si existe comunicación con un dispositivo Modbus RTU.
    
    Args:
        port (str): Puerto COM del convertidor RS485
        baudrate (int): Velocidad de comunicación
        parity (str): Paridad ('N' para ninguna, 'E' para par, 'O' para impar)
        stopbits (int): Bits de parada
        bytesize (int): Tamaño de byte
        timeout (int): Tiempo de espera para la respuesta
    """
    print(f"Iniciando prueba de comunicación Modbus RTU en {port}")
    print(f"Configuración: {baudrate} baudios, {bytesize}{parity}{stopbits}, timeout {timeout}s")
    
    # Crear cliente Modbus RTU
    client = ModbusSerialClient(
        port=port,
        baudrate=baudrate,
        parity=parity,
        stopbits=stopbits,
        bytesize=bytesize,
        timeout=timeout
    )
    
    # Intentar conectar
    connection = client.connect()
    if connection:
        print(f"Conexión establecida con el puerto {port}")
        
        # Dirección del dispositivo Danfoss
        slave_address = 0
        
        print(f"Probando comunicación con dispositivo en dirección {slave_address}...")
        
        try:
            # Intentar leer algunos registros básicos para verificar comunicación
            test_registers = [
                (0, 1),     # Primer registro
                (1, 1),     # Segundo registro
                (100, 1),   # Registro 100
                (400, 1)    # Registro 400
            ]
            
            communication_successful = False
            
            for start_register, count in test_registers:
                print(f"Probando lectura del registro {start_register}...")
                
                # Intentar leer registro de retención
                response = client.read_holding_registers(start_register, count, slave=slave_address)
                if not hasattr(response, 'isError') or not response.isError():
                    print(f"¡COMUNICACIÓN EXITOSA! Registro {start_register} leído: {response.registers}")
                    communication_successful = True
                    break
                
                # Intentar leer registro de entrada
                response = client.read_input_registers(start_register, count, slave=slave_address)
                if not hasattr(response, 'isError') or not response.isError():
                    print(f"¡COMUNICACIÓN EXITOSA! Registro de entrada {start_register} leído: {response.registers}")
                    communication_successful = True
                    break
                
                time.sleep(0.5)
            
            if communication_successful:
                print("\n✅ COMUNICACIÓN ESTABLECIDA CON ÉXITO ✅")
                print("La conexión con el dispositivo Danfoss MCX06D está funcionando correctamente.")
            else:
                print("\n❌ NO SE PUDO ESTABLECER COMUNICACIÓN ❌")
                print("Sugerencias:")
                print("1. Verifica que la dirección del esclavo sea 1")
                print("2. Confirma que los parámetros de comunicación sean correctos (38400, E, 1)")
                print("3. Revisa las conexiones físicas del convertidor RS485")
        
        except ModbusException as e:
            print(f"\n❌ ERROR DE COMUNICACIÓN MODBUS: {e}")
        
        except Exception as e:
            print(f"\n❌ ERROR GENERAL: {e}")
        
        finally:
            client.close()
            print("Conexión cerrada")
    else:
        print(f"\n❌ NO SE PUDO ESTABLECER CONEXIÓN CON EL PUERTO {port}")
        print("Verifica que:")
        print("1. El puerto COM3 esté disponible")
        print("2. El convertidor USB-RS485 esté correctamente instalado")
        print("3. No haya otro programa utilizando el puerto")

def read_temperature_register(port='COM3', baudrate=9600, parity='N', stopbits=1, bytesize=8, timeout=1, slave_address=0):
    """
    Lee específicamente el registro de temperatura en la dirección 0x0000 (0 decimal)
    e intenta interpretarlo de diferentes maneras.
    
    Args:
        port (str): Puerto COM del convertidor RS485
        baudrate (int): Velocidad de comunicación
        parity (str): Paridad ('N' para ninguna, 'E' para par, 'O' para impar)
        stopbits (int): Bits de parada
        bytesize (int): Tamaño de byte
        timeout (int): Tiempo de espera para la respuesta
        slave_address (int): Dirección del esclavo Modbus
    """
    print(f"Leyendo registro de temperatura en dirección 0x0000 (0 decimal)")
    print(f"Configuración: {baudrate} baudios, {bytesize}{parity}{stopbits}, timeout {timeout}s")
    print(f"Dirección del esclavo: {slave_address}")
    
    # Crear cliente Modbus RTU
    client = ModbusSerialClient(
        port=port,
        baudrate=baudrate,
        parity=parity,
        stopbits=stopbits,
        bytesize=bytesize,
        timeout=timeout
    )
    
    # Intentar conectar
    connection = client.connect()
    if connection:
        print(f"Conexión establecida con el puerto {port}")
        
        try:
            # Probar diferentes tipos de registros y métodos de lectura
            print("\n--- Intentando diferentes métodos de lectura para el registro 0x0000 ---")
            
            # 1. Leer como holding register (función 03)
            print("\nProbando como holding register (función 03):")
            try:
                response = client.read_holding_registers(0, 1, slave=slave_address)
                if not hasattr(response, 'isError') or not response.isError():
                    value = response.registers[0]
                    print(f"Valor leído: {value} (decimal) / 0x{value:04X} (hex)")
                    print(f"Interpretado como temperatura (directo): {value}°C")
                    print(f"Interpretado como temperatura (÷10): {value/10.0}°C")
                    print(f"Interpretado como temperatura (÷100): {value/100.0}°C")
                    
                    # Interpretar como valor con signo (complemento a 2)
                    if value > 32767:
                        signed_value = value - 65536
                    else:
                        signed_value = value
                    print(f"Interpretado como temperatura con signo: {signed_value}°C")
                    print(f"Interpretado como temperatura con signo (÷10): {signed_value/10.0}°C")
                else:
                    print(f"Error al leer holding register: {response}")
            except Exception as e:
                print(f"Error al leer holding register: {e}")
            
            # 2. Leer como input register (función 04)
            print("\nProbando como input register (función 04):")
            try:
                response = client.read_input_registers(0, 1, slave=slave_address)
                if not hasattr(response, 'isError') or not response.isError():
                    value = response.registers[0]
                    print(f"Valor leído: {value} (decimal) / 0x{value:04X} (hex)")
                    print(f"Interpretado como temperatura (directo): {value}°C")
                    print(f"Interpretado como temperatura (÷10): {value/10.0}°C")
                    print(f"Interpretado como temperatura (÷100): {value/100.0}°C")
                    
                    # Interpretar como valor con signo (complemento a 2)
                    if value > 32767:
                        signed_value = value - 65536
                    else:
                        signed_value = value
                    print(f"Interpretado como temperatura con signo: {signed_value}°C")
                    print(f"Interpretado como temperatura con signo (÷10): {signed_value/10.0}°C")
                else:
                    print(f"Error al leer input register: {response}")
            except Exception as e:
                print(f"Error al leer input register: {e}")
            
            # 3. Probar leyendo 2 registros (por si la temperatura es un valor de 32 bits)
            print("\nProbando lectura de 2 registros (por si es valor de 32 bits):")
            try:
                response = client.read_holding_registers(0, 2, slave=slave_address)
                if not hasattr(response, 'isError') or not response.isError():
                    value1 = response.registers[0]
                    value2 = response.registers[1]
                    print(f"Valores leídos: {value1}, {value2}")
                    
                    # Combinar como valor de 32 bits (big endian)
                    combined_be = (value1 << 16) | value2
                    print(f"Combinado (big endian): {combined_be}")
                    print(f"Interpretado como temperatura (÷10): {combined_be/10.0}°C")
                    
                    # Combinar como valor de 32 bits (little endian)
                    combined_le = (value2 << 16) | value1
                    print(f"Combinado (little endian): {combined_le}")
                    print(f"Interpretado como temperatura (÷10): {combined_le/10.0}°C")
                else:
                    print(f"Error al leer 2 registros: {response}")
            except Exception as e:
                print(f"Error al leer 2 registros: {e}")
            
            # 4. Probar con diferentes direcciones cercanas
            print("\nProbando direcciones cercanas:")
            for addr in range(1, 6):  # Probar registros 1-5
                try:
                    print(f"\nProbando registro {addr} (0x{addr:04X}):")
                    response = client.read_holding_registers(addr, 1, slave=slave_address)
                    if not hasattr(response, 'isError') or not response.isError():
                        value = response.registers[0]
                        print(f"Valor leído: {value} (decimal) / 0x{value:04X} (hex)")
                        print(f"Interpretado como temperatura (directo): {value}°C")
                        print(f"Interpretado como temperatura (÷10): {value/10.0}°C")
                    else:
                        print(f"Error al leer registro {addr}: {response}")
                except Exception as e:
                    print(f"Error al leer registro {addr}: {e}")
        
        except ModbusException as e:
            print(f"\nError de comunicación Modbus: {e}")
        
        except Exception as e:
            print(f"\nError general: {e}")
        
        finally:
            client.close()
            print("\nConexión cerrada")
    else:
        print(f"\nNo se pudo establecer conexión con el puerto {port}")
        print("Verifica que:")
        print("1. El puerto COM3 esté disponible")
        print("2. El convertidor USB-RS485 esté correctamente instalado")
        print("3. No haya otro programa utilizando el puerto")

def monitor_temperature(port='COM3', baudrate=9600, parity='N', stopbits=1, bytesize=8, timeout=1, 
                       slave_address=0, register=0, register_type='holding', scale=1.0):
    """
    Monitorea continuamente el registro de temperatura.
    
    Args:
        port (str): Puerto COM del convertidor RS485
        baudrate (int): Velocidad de comunicación
        parity (str): Paridad ('N' para ninguna, 'E' para par, 'O' para impar)
        stopbits (int): Bits de parada
        bytesize (int): Tamaño de byte
        timeout (int): Tiempo de espera para la respuesta
        slave_address (int): Dirección del esclavo Modbus
        register (int): Registro a monitorear
        register_type (str): Tipo de registro ('holding' o 'input')
        scale (float): Factor de escala para el valor leído
    """
    print(f"Monitoreando temperatura en registro {register_type} {register} (0x{register:04X})")
    print(f"Configuración: {baudrate} baudios, {bytesize}{parity}{stopbits}, escala {scale}")
    print(f"Dirección del esclavo: {slave_address}")
    print("Presiona Ctrl+C para detener el monitoreo")
    
    # Crear cliente Modbus RTU
    client = ModbusSerialClient(
        port=port,
        baudrate=baudrate,
        parity=parity,
        stopbits=stopbits,
        bytesize=bytesize,
        timeout=timeout
    )
    
    # Intentar conectar
    connection = client.connect()
    if connection:
        print(f"Conexión establecida con el puerto {port}")
        
        try:
            while True:
                # Leer registro según el tipo especificado
                if register_type.lower() == 'holding':
                    response = client.read_holding_registers(register, 1, slave=slave_address)
                elif register_type.lower() == 'input':
                    response = client.read_input_registers(register, 1, slave=slave_address)
                else:
                    print(f"Tipo de registro no válido: {register_type}")
                    break
                
                if not hasattr(response, 'isError') or not response.isError():
                    value = response.registers[0]
                    scaled_value = value * scale
                    
                    # Mostrar valor con formato según la escala
                    if scale == 1:
                        print(f"[{time.strftime('%H:%M:%S')}] Temperatura: {scaled_value}°C")
                    elif scale == 0.1:
                        print(f"[{time.strftime('%H:%M:%S')}] Temperatura: {scaled_value:.1f}°C")
                    elif scale == 0.01:
                        print(f"[{time.strftime('%H:%M:%S')}] Temperatura: {scaled_value:.2f}°C")
                    else:
                        print(f"[{time.strftime('%H:%M:%S')}] Temperatura: {scaled_value}°C")
                else:
                    print(f"[{time.strftime('%H:%M:%S')}] Error al leer registro: {response}")
                
                time.sleep(1)  # Actualizar cada segundo
        
        except KeyboardInterrupt:
            print("\nMonitoreo detenido por el usuario")
        
        except ModbusException as e:
            print(f"\nError de comunicación Modbus: {e}")
        
        except Exception as e:
            print(f"\nError general: {e}")
        
        finally:
            client.close()
            print("\nConexión cerrada")
    else:
        print(f"\nNo se pudo establecer conexión con el puerto {port}")
        print("Verifica que:")
        print("1. El puerto esté disponible")
        print("2. El convertidor USB-RS485 esté correctamente instalado")
        print("3. No haya otro programa utilizando el puerto")

if __name__ == "__main__":
    # Ejecutar la función de lectura específica del registro de temperatura
    """read_temperature_register(
        port='COM3',
        baudrate=9600,
        parity='N',
        stopbits=1,
        bytesize=8,
        timeout=1,
        slave_address=0
    )
    
    # Descomentar para monitorear continuamente la temperatura
    # (ajustar los parámetros según los resultados de la lectura inicial)
"""
    monitor_temperature(
        port='COM3',
        baudrate=9600,
        parity='N',
        stopbits=1,
        bytesize=8,
        timeout=1,
        slave_address=0,
        register=0,        # Ajustar según el registro correcto
        register_type='holding',  # 'holding' o 'input'
        scale=0.1          # Ajustar según el factor de escala correcto
    )
