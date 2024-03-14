import os
import inspect
import argparse
import plugins.base

parser = argparse.ArgumentParser()
parser.add_argument('-p', '--plugin_name', dest='pluginName', required=True,
                    help="Имя плагина (Имя py-файла без расширения)")
parser.add_argument('-i', '--input', dest='pluginInput', help="Папка с входными файлами плагина")
parser.add_argument('-o', '--output', dest='pluginOutput', help="Папка для вывода файлов плагина")

args = parser.parse_args()

plugin_dir = "macros/plugins"

# Сюда добавляем имена загруженных модулей
modules = []

# Перебирем файлы в папке plugins
print("11")
for fname in os.listdir(plugin_dir):
    print("12")
    # Нас интересуют только файлы с расширением .py
    if fname.endswith(".py"):
        # Обрежем расширение .py у имени файла
        module_name = fname[: -3]
        print(module_name)
        # Пропустим файлы base.py и __init__.py
        if module_name != "base" and module_name != "__init__":
            # Загружаем модуль и добавляем его имя в список загруженных модулей
            package_obj = __import__("plugins." + module_name)
            modules.append(module_name)

# Перебираем загруженные модули
for modulename in modules:
    if modulename == args.pluginName:

        module_obj = getattr(package_obj, modulename)
        # Перебираем все, что есть внутри модуля
        for elem in dir(module_obj):
            obj = getattr(module_obj, elem)
            # Это класс?
            if inspect.isclass(obj):
                # Класс производный от baseplugin?
                if issubclass(obj, plugins.base.basePlugin):
                    # Создаем экземпляр и выполняем функцию run
                    a = obj()
                    a.run(args.pluginInput, args.pluginOutput)
