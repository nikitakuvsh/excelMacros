import pyfiglet
import os
import questionary


def input_file():
    text = 'BONUS CALC'
    ascii_art = pyfiglet.figlet_format(text)
    print(ascii_art)
    print("v 1.0.0")

    print('Чтобы скрипт всё рассчитал, файл xlsx должен лежать в одной папке с этим скриптом!')

    files = [
        f for f in os.listdir('.')
        if f.endswith('.xlsx') and 'generated' not in f
    ]

    if not files:
        print("❌ В папке нет .xlsx файлов!")
        return None

    if len(files) == 1:
        file_path = files[0]
        print(f"Найден единственный файл: {file_path}")
        return file_path

    # ✔️ если несколько — показываем выбор
    file_path = questionary.select(
        "Выберите Excel файл:",
        choices=files
    ).ask()

    return file_path