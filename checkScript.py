#Autor: Navin Ruas
import threading
import psutil
import os
import time
from Control import ochamado

def check_script():
    while True:
        # Pegar o PID do processo
        pid = os.getpid()
        # Verificar se o processo está rodando
        if not psutil.pid_exists(pid):
            # Se não estiver rodando, executar o script
            ochamado()
        # Dorme por 1 minuto para reduzir o uso de memória
        time.sleep(60)

if __name__ == '__main__':
    # Cria uma thread para executar a função check_script em segundo plano
    t = threading.Thread(target=check_script)
    t.start()