#немного кода для тестов
import sys
import venv
import os
import subprocess
import shutil

def create_venv():
    venv_dir = "venv"
    
    # Проверка существования папки venv
    if os.path.exists(venv_dir):
        print(f"Папка '{venv_dir}' уже существует. Удалить её? (y/n)")
        choice = input().strip().lower()
        if choice == 'y':
            shutil.rmtree(venv_dir)
        else:
            print("Используем существующую виртуальную среду.")
            return venv_dir
    
    # Создание виртуальной среды
    venv.create(venv_dir, with_pip=True)
    print(f"Виртуальная среда создана в '{venv_dir}'")
    return venv_dir

def activate_venv(venv_dir ,manage_dir):
    # Команда активации в зависимости от ОС
    if sys.platform == "win32":
        activate_path = os.path.join(venv_dir, "Scripts", "activate.bat")
        python = os.path.join(venv_dir, "Scripts", "python.exe")
        manage_dir = os.path.join(manage_dir, 'manage.py')
        
        command = f'cmd.exe /K "{activate_path} &&  {python} {manage_dir} runserver"'
    else:
        activate_path = os.path.join(venv_dir, "bin", "activate")
        command = f'bash -c "source {activate_path} && exec bash"'

    # Запуск нового терминала с активированной средой
    try:
        subprocess.Popen(command, shell=True)
        print(u"Запущен новый терминал с активированной виртуальной средой.")
    except Exception as e:
        print(f"Ошибка: {str(e)}")

if __name__ == "__main__":
    venv_dir = r"venv"
    manage_dir = r"ProjectDir"
    activate_venv(venv_dir, manage_dir)
