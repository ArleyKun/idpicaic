import psutil
import time

print("closing all word files...")

while True:
    word_processes = [proc for proc in psutil.process_iter(['name']) if proc.info['name'] == 'WINWORD.EXE']

    if not word_processes:
        print("no more word processes found. Script stopped.")
        break

    for proc in word_processes:
        try:
            proc.kill()
            print(f"Killed process: {proc.pid}")
        except Exception as e:
            print(f"Error killing process: {e}")

    time.sleep(0.5)
