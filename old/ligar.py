import schedule
import os

def ligar_computador():
    os.system("shutdown /s /t 1")

schedule.every().day.at("03:55").do(ligar_computador)

while True:
    schedule.run_pending()