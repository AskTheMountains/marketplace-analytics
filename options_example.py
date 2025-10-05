import os
import sys
import json

# 0 - Example Client
client_number = 0  # <--- Вручную менять здесь


def loadFromJSON(file_name):
    with open(file_name, "r", encoding="utf8") as f:
        return json.load(f)

settings = loadFromJSON('settings.json')


# Поиск переменных окружения и парметров командной строки (они в приоритете)
env_val = os.environ.get('marketplace_client')
arg_val = None

for i, a in enumerate(sys.argv):
    if a in ("--client", "-c") and i + 1 < len(sys.argv):
        arg_val = sys.argv[i + 1]
        break

if env_val is not None:
    client_number = int(env_val)
elif arg_val is not None:
    client_number = int(arg_val)
# Если ни переменной среды, ни аргумента — используется значение, которое вы поставили выше
