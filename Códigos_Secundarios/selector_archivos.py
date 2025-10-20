import os
import shutil
import webview

base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

class Api:
    def __init__(self):
        self.word_path = None
        self.excel_path = None

    def elegirWord(self):
        """Abrir diálogo de selección para Word y copiar a carpeta base"""
        file_types = ('Archivos Word (*.docx)',)
        result = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG, file_types=file_types
        )
        if result:
            path = result[0]
            destino = os.path.join(base_dir, os.path.basename(path))
            shutil.copy2(path, destino)
            self.word_path = destino
            print(f"[OK] Word copiado: {destino}")
            return {"status": "ok", "path": destino}
        return {"status": "error", "msg": "No se seleccionó archivo Word"}

    def elegirExcel(self):
        """Abrir diálogo de selección para Excel y copiar a carpeta base"""
        file_types = ('Archivos Excel (*.xlsm;*.xlsx;*.xls)',)
        result = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG, file_types=file_types
        )
        if result:
            path = result[0]
            destino = os.path.join(base_dir, os.path.basename(path))
            shutil.copy2(path, destino)
            self.excel_path = destino
            print(f"[OK] Excel copiado: {destino}")
            return {"status": "ok", "path": destino}
        return {"status": "error", "msg": "No se seleccionó archivo Excel"}

    def continuar(self):
        """Cerrar ventana solo si ambos archivos están cargados"""
        if self.word_path and self.excel_path:
            print("[OK] Ambos archivos listos, cerrando ventana.")
            webview.windows[0].destroy()
            return {"word": self.word_path, "excel": self.excel_path}
        else:
            print("[ERROR] Falta seleccionar Word o Excel.")
            return {"status": "error", "msg": "Debe seleccionar Word y Excel antes de continuar"}


def seleccionar_archivos():
    api = Api()
    html_path = os.path.join(base_dir, "Recursos_UI", "selector.html")

    window = webview.create_window(
        "Seleccionar archivos",
        url=f"file://{html_path}",
        width=600,
        height=550,
        js_api=api,
        frameless=False
    )

    webview.start()
    return {"word": api.word_path, "excel": api.excel_path}


if __name__ == "__main__":
    print(seleccionar_archivos())

