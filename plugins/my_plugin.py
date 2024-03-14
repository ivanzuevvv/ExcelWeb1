import plugins.base

class pluginClass(plugins.base.basePlugin):
    def __init__(self):
        pass
    def run(self, pluginInput, pluginOutput):
        print("Запущен пробный скрипт №001.")
        print("pluginInput=", pluginInput)
        print("pluginOutput=", pluginOutput)