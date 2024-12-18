import datetime as dt

class Logger:
    def __init__(self, log_file):
        self.log_file = log_file

    def log(self, message):
        try:
            timestamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(self.log_file, mode='a', encoding='utf8') as fp:
                fp.write(f"{timestamp} - {message}\n")
            print(f"{timestamp} - {message}")
        except Exception as e:
            print(f"Error en el logger{e}")
