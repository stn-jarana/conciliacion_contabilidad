from pywinauto import Application, findwindows, Desktop
from pathlib import Path
import time
import os
import sys
import dotenv

dotenv.load_dotenv()

usuario = os.getenv('usuario', '').upper()
clave = os.getenv('clave', '')
ruta_exe = os.getenv('ruta_exe')
directorio_trabajo = os.getenv('directorio_trabajo')

def abrir_app() -> Application:
    app = Application(backend='uia').start(ruta_exe, work_dir=directorio_trabajo)
    print('Abriendo contanet')

    time.sleep(1)

    window = app.top_window()
    window.set_focus()

    time.sleep(1)

    window.type_keys(usuario, with_spaces=True, pause=0.01, vk_packet=True)
    window.type_keys("{TAB}", vk_packet=False)
    window.type_keys(clave, with_spaces=True, pause=0.01, vk_packet=True)
    window.type_keys("{ENTER}", vk_packet=False)

    return app


def seleccionar_empresa(app: Application, ruc_cliente: str='20376729126', anio='2024'):
    window = app.top_window()
    window.set_focus()

    empresa = window.child_window(auto_id="txtBusqEmp", control_type="Edit")
    empresa.type_keys(ruc_cliente, with_spaces=True, pause=0.01, vk_packet=True)
    empresa.type_keys("{TAB}", vk_packet=False)

    ejercicio = window.child_window(title=anio, control_type="ListItem")
    ejercicio.type_keys("{ENTER}", vk_packet=False)


def seleccionar_import(app: Application):
    window = app.top_window()
    window.set_focus()

    for btn in window.descendants(control_type="Button"):
        texto = btn.window_text()

        if texto == "IMPORTACIÓN DE DOCUMENTOS":
            print(btn)

            btn.invoke()

            return


def elegir_procesos(app: Application):
    window = app.top_window()
    window.set_focus()

    try:
        procesos_item = window.child_window(title="Nodo1", control_type="TreeItem")
        procesos_item.wait('ready')

        try:
            procesos_item.double_click_input()
            procesos_item.type_keys("{DOWN}{DOWN}{ENTER}")
        except Exception as e:
            procesos_item.expand()


    except Exception as e:
        print(f"Error al intentar abrir Procesos: {e}")


def import_conta(app, archivo='Plantilla comisiones 06-2026 Bcp Dolares.xlsm', eliminar_archivo: bool = False):
    window = app.top_window()
    window.set_focus()

    cargar = window.child_window(title="Cargar", control_type="Button")
    cargar.click_input()

    time.sleep(1)

    if isinstance(archivo, Path):
        ruta_archivo = archivo.resolve()
    else:
        p = Path(archivo)
        ruta_archivo = p.resolve() if p.is_absolute() else (Path(__file__).parent / p).resolve()

    if not ruta_archivo.exists():
        app.kill()
        raise FileNotFoundError(f'No se encontró el archivo: {ruta_archivo}')

    for _ in range(20):
        elementos = findwindows.find_elements(
            backend="win32", class_name="#32770", title="Abrir",
            top_level_only=True,
        )
        if elementos:
            break
        time.sleep(0.5)
    else:
        raise RuntimeError("No apareció el diálogo 'Abrir'")

    dlg = Desktop(backend="win32").window(handle=elementos[-1].handle)

    combo = dlg.child_window(class_name="ComboBox", found_index=0)
    edit = combo.child_window(class_name="Edit")
    edit.wait('ready', timeout=5)

    edit.set_edit_text(str(ruta_archivo))
    time.sleep(0.2)
    edit.type_keys("{ENTER}")

    time.sleep(2)

    validar = window.child_window(title="Validar", control_type="Button")
    validar.click_input()

    time.sleep(3)

    mensaje = window.child_window(title="Mensaje Sistema", control_type="Window")
    mensaje.wait('ready', timeout=5)

    aceptar = mensaje.child_window(title="Aceptar", control_type="Button")
    aceptar.wait('ready', timeout=5)
    aceptar.click_input()

    time.sleep(2)

    importar = window.child_window(title="Importar", control_type="Button")
    importar.click_input()

    time.sleep(2)

    mensaje = window.child_window(title="Mensaje Sistema", control_type="Window")
    mensaje.wait('ready', timeout=5)

    aceptar = mensaje.child_window(title="Aceptar", control_type="Button")
    aceptar.wait('ready', timeout=5)
    aceptar.click_input()

    time.sleep(2)

    if eliminar_archivo:
        ruta_archivo.unlink(missing_ok=True)
        print('Archivo eliminado:', ruta_archivo)

    app.kill()


def ingresar_asiento_contanet(
    ruta_archivo: Path | str = 'Plantilla comisiones 06-2026 Bcp Dolares.xlsm',
    ruc_cliente: str = '20376729126',
    anio: str = '2024',
    eliminar_archivo: bool = False,
) -> bool:
    """Ejecuta el ingreso automatizado del asiento en Contanet de forma independiente.

    Parameters
    ----------
    ruta_archivo : Path | str
        Ruta del archivo Excel (.xlsm) con el asiento a importar.
    ruc_cliente : str
        RUC de la empresa en Contanet.
    anio : str
        Año del ejercicio contable.
    eliminar_archivo : bool
        Si es True, elimina el archivo tras importarlo con éxito. Por defecto False.

    Returns
    -------
    bool
        True si el proceso se completó con éxito.
    """
    app = abrir_app()
    try:
        seleccionar_empresa(app, ruc_cliente=ruc_cliente, anio=str(anio))
        seleccionar_import(app)
        elegir_procesos(app)
        import_conta(app, archivo=ruta_archivo, eliminar_archivo=eliminar_archivo)
        return True
    except Exception as e:
        try:
            app.kill()
        except Exception:
            pass
        raise e


if __name__ == '__main__':
    # Ejecución manual y separada del flujo general
    print("Iniciando ingreso en Contanet de manera independiente...")
    ingresar_asiento_contanet()

