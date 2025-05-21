import socket

# 設備參數
HOST = '192.168.1.2'  # 電力計 IP
PORT = 3300           # 通訊 Port（LAN 通訊為 3300）
COMMAND = ":MEAS? U,I,P,WH\n"  # SCPI 查詢指令

def parse_measurement(value_str):
    # Remove the measurement type (U, I, P, etc.) and convert to float
    return float(value_str.split()[1])

# 建立 socket 並傳送指令
try:
    with socket.create_connection((HOST, PORT), timeout=5) as sock:
        print("Connected to PW3335.")
        sock.sendall(COMMAND.encode('ascii'))  # 傳送查詢指令
        
        response = sock.recv(1024).decode('ascii').strip()  # 接收回應
        measurements = response.split(';')
        #print(response)
        # Parse each measurement
        for measure in measurements:
            if measure.startswith('U'):
                voltage = parse_measurement(measure)
                print(f"Voltage: {voltage} V")
            elif measure.startswith('I'):
                current = parse_measurement(measure)
                print(f"Current: {current} A")
            elif measure.startswith('P'):
                power = parse_measurement(measure)
                print(f"Power: {power} W")
            elif measure.startswith('S'):
                apparent_power = parse_measurement(measure)
                print(f"Apparent Power: {apparent_power} VA")
            elif measure.startswith('Q'):
                reactive_power = parse_measurement(measure)
                print(f"Reactive Power: {reactive_power} VAR")
            elif measure.startswith('PF'):
                power_factor = parse_measurement(measure)
                print(f"Power Factor: {power_factor}")
            elif measure.startswith('DEGAC'):
                phase_angle = parse_measurement(measure)
                print(f"Phase Angle: {phase_angle} degrees")
            elif measure.startswith('FREQU'):
                frequency = parse_measurement(measure)
                print(f"Frequency: {frequency} Hz")
            elif measure.startswith('WP'):
                wp = parse_measurement(measure)
                print(f"WP: {wp} Wh")

except socket.timeout:
    print("Connection timed out. Please check IP and port.")
except ConnectionRefusedError:
    print("Connection refused. Is the device powered on and reachable?")
except Exception as e:
    print("Error:", e)
