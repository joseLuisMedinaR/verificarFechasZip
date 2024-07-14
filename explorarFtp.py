import os
import zipfile
import pandas as pd
from datetime import datetime, timedelta
import flet as ft
from ftplib import FTP
from config import FTP_CONFIG

# Función para obtener información de archivos dentro de archivos ZIP en una carpeta raíz dada
def get_files_info_ftp(ftp, root_path, progress_callback, connection_type):
    data = []
    dir_list = []
    ftp.dir(dir_list.append)

    if connection_type == "Digitar":
        for item in dir_list:
            words = item.split()
            if "<DIR>" in words:
                folder_name = words[-1]
                try:
                    ftp.cwd(folder_name)
                    files = ftp.nlst()
                    if "BrioWeb.zip" in files:
                        zip_date = datetime.now().strftime('%Y-%m-%d')
                        with open("BrioWeb.zip", 'wb') as f:
                            ftp.retrbinary(f'RETR BrioWeb.zip', f.write)
                        with zipfile.ZipFile("BrioWeb.zip", 'r') as zip_ref:
                            for file_info in zip_ref.infolist():
                                file_date = datetime(*file_info.date_time).strftime('%Y-%m-%d')
                                file_hour = datetime(*file_info.date_time).strftime('%H:%M:%S')
                                data.append([folder_name, "BrioWeb.zip", zip_date, file_date, file_info.filename, file_hour])
                        os.remove("BrioWeb.zip")
                    ftp.cwd('..')
                except Exception as e:
                    print(f"No se pudo acceder a la carpeta '{folder_name}': {str(e)}")
                progress_callback(f"Verificando: {folder_name}")
    else:
        for item in dir_list:
            words = item.split()
            if words[0].startswith('d'):
                folder_name = words[-1]
                try:
                    ftp.cwd(folder_name)
                    files = ftp.nlst()
                    if "BrioWeb.zip" in files:
                        zip_date = datetime.now().strftime('%Y-%m-%d')
                        with open("BrioWeb.zip", 'wb') as f:
                            ftp.retrbinary(f'RETR BrioWeb.zip', f.write)
                        with zipfile.ZipFile("BrioWeb.zip", 'r') as zip_ref:
                            for file_info in zip_ref.infolist():
                                file_date = datetime(*file_info.date_time).strftime('%Y-%m-%d')
                                file_hour = datetime(*file_info.date_time).strftime('%H:%M:%S')
                                data.append([folder_name, "BrioWeb.zip", zip_date, file_date, file_info.filename, file_hour])
                    ftp.cwd('..')
                except Exception as e:
                    print(f"No se pudo acceder a la carpeta '{folder_name}': {str(e)}")
                progress_callback(f"Verificando: {folder_name}")

    return data

# Función para guardar los datos en un archivo de Excel
def save_to_excel(data, save_path):
    if not save_path.endswith('.xlsx'):
        save_path += '.xlsx'
    df = pd.DataFrame(data, columns=['Carpeta', 'Archivo Zip', 'Fecha Archivo Zip', 'Fecha Archivo en Zip', 'Archivo en Zip', 'Hora Archivo'])
    df.to_excel(save_path, index=False, engine='openpyxl')
    return save_path

# Función para cargar datos desde un archivo de Excel y convertirlos en filas de DataRow de Flet
def load_data_from_excel(excel_path):
    df = pd.read_excel(excel_path, engine='openpyxl')
    rows = []
    for index, row in df.iterrows():
        cells = []
        for col_name, cell in row.items():
            # Determinar el color de texto basado en la comparación de fechas
            #text_color = ft.colors.BLUE_GREY_800 if row['Fecha Archivo Zip'] == row['Fecha Archivo en Zip'] else ft.colors.CYAN_50
            text_color = ft.colors.CYAN_50 if row['Fecha Archivo en Zip'] < (datetime.now()+timedelta(days=-1)).strftime('%Y-%m-%d') else ft.colors.BLUE_GREY_800
            # Estilo de texto con el color determinado
            text_style = ft.TextStyle(color=text_color)
            # Crear el objeto Text con el estilo y texto
            text = ft.Text(str(cell), style=text_style)
            # Crear la celda de datos con el texto
            cells.append(ft.DataCell(text))
        
        # Determinar el color de fondo de la fila
        #row_color = ft.colors.RED_ACCENT_400 if row['Fecha Archivo Zip'] != row['Fecha Archivo en Zip'] else ft.colors.LIGHT_BLUE_200
        row_color = ft.colors.RED_ACCENT_400 if row['Fecha Archivo en Zip'] < (datetime.now()+timedelta(days=-1)).strftime('%Y-%m-%d') else ft.colors.LIGHT_BLUE_200
        
        # Agregar la fila de datos con las celdas y el color de fondo
        rows.append(ft.DataRow(cells, color=row_color))

    return df, rows

def main(page: ft.Page):
    page.title = "Verificar Fechas Archivos Zip - by JoseLu - 2024"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.bgcolor = "#00bfff"
    page.window.maximizable = True
    page.window.minimizable = True
    page.scroll = True
    page.window.maximized = True

    path_var = ft.TextField(label="Ruta de las carpetas", width=400, disabled=True)
    result_label = ft.Text(value="", color=ft.colors.WHITE)
    connection_status = ft.Text(value="", color=ft.colors.TEAL_ACCENT_100)

    df = pd.DataFrame()
    sort_state = {'Carpeta': True, 'Archivo Zip': True, 'Fecha Archivo Zip': True, 'Fecha Archivo en Zip': True, 'Archivo en Zip': True, 'Hora Archivo': True}

    data_table = ft.DataTable(
        bgcolor="white",
        border=ft.border.all(2, ft.colors.PURPLE_300),
        border_radius=10,
        vertical_lines=ft.BorderSide(3, ft.colors.CYAN_200),
        horizontal_lines=ft.BorderSide(1, ft.colors.BLUE_100),
        sort_column_index=0,
        sort_ascending=True,
        heading_row_color=ft.colors.BLACK12,
        heading_row_height=100,
        data_row_color={"hovered": "0x30FF0000"},
        divider_thickness=0,
        columns=[
            ft.DataColumn(ft.Text("Carpeta"), on_sort=lambda e: sort_table('Carpeta', xlsx_path)),
            ft.DataColumn(ft.Text("Archivo Zip"), on_sort=lambda e: sort_table('Archivo Zip', xlsx_path)),
            ft.DataColumn(ft.Text("Fecha Archivo Zip"), on_sort=lambda e: sort_table('Fecha Archivo Zip', xlsx_path)),
            ft.DataColumn(ft.Text("Fecha Archivo en Zip"), on_sort=lambda e: sort_table('Fecha Archivo en Zip', xlsx_path)),
            ft.DataColumn(ft.Text("Archivo en Zip"), on_sort=lambda e: sort_table('Archivo en Zip', xlsx_path)),
            ft.DataColumn(ft.Text("Hora Archivo"), on_sort=lambda e: sort_table('Hora Archivo', xlsx_path)),
        ],
        rows=[],
    )

    # Definir xlsx_path en el ámbito global de main
    xlsx_path = ""
    ftp = None
    ftp_config = None

    # Función que verifica y guarda los datos en un archivo de Excel
    def verify_and_save(save_path):
        nonlocal xlsx_path, ftp, ftp_config # Asegurar que xlsx_path sea modificable dentro de la función
        root_path = path_var.value
        if root_path:
            result_label.value = "Estamos verificando cada carpeta en el FTP, aguarde un momento por favor."
            page.update()

            try:
                if ftp:
                    # Obtener información de los archivos en el FTP
                    files_info = get_files_info_ftp(ftp, ftp.pwd(), lambda msg: update_progress(msg), dd.value)
                    
                    # Guardar la información en un archivo Excel
                    xlsx_path = save_to_excel(files_info, save_path)
                    
                    if xlsx_path:
                        result_label.value = f"Información guardada en:\n{xlsx_path}"
                        nonlocal df
                        df, data_table.rows = load_data_from_excel(xlsx_path)
                        page.update()
                    else:
                        result_label.value = "No se ha seleccionado una ubicación para guardar."
                    page.update()
                else:
                    raise ValueError("No se ha establecido una conexión FTP válida.")

            except Exception as e:
                result_label.value = f"No se pudo verificar y guardar: {str(e)}"
                page.update()

        else:
            open_dlg_modal_help()

    # Función para actualizar el progreso
    def update_progress(msg):
        result_label.value = msg
        page.update()

    # Función para ordenar la tabla por una columna específica
    def sort_table(column_name, excel_path):
        nonlocal df
        ascending = not sort_state[column_name]
        df, _ = load_data_from_excel(excel_path)
        df = df.sort_values(by=column_name, ascending=ascending)
        df.to_excel(excel_path, index=False, engine='openpyxl')
        _, data_table.rows = load_data_from_excel(excel_path)
        sort_state[column_name] = ascending
        page.update()

    def accept(e):
        nonlocal xlsx_path, ftp, ftp_config
        if not path_var.value:
            open_dlg_modal_help()
        elif not ftp:
            connection_status.value = "Debe seleccionar una conexión FTP antes de continuar."
            page.update()
        else:
            verify_and_save(xlsx_path)

    #Función que guarda la ruta del archivo
    def on_file_picker_result(e):
        if e.path:
            path_var.value = e.path
            verify_and_save(e.path)
            page.update()

    # Función para cerrar el cuadro de diálogo de ayuda
    def close_dlg_help(e):
        dlg_modal_help.open = False
        page.update()

    # Cuadro de diálogo de ayuda
    dlg_modal_help = ft.AlertDialog(
        modal=True,
        title=ft.Text("Aviso"),
        content=ft.Text(
            "Debe seleccionar una carpeta antes de continuar."
        ),
        actions=[
            ft.TextButton("Cerrar", on_click=close_dlg_help),
        ],
        actions_alignment=ft.MainAxisAlignment.END,
    )

    # Función para abrir el cuadro de diálogo de ayuda
    def open_dlg_modal_help():
        dlg_modal_help.open = True
        page.update()

    # Función que toma el clic en el combo de conexiones
    def dropdown_changed(e):
        nonlocal ftp, ftp_config  # Declarar variables como no locales
        page.update()
        if dd.value == "Movilsol Oficina":
            ftp_config = FTP_CONFIG['MovilsolOficina']
        elif dd.value == "Movilsol desde Afuera":
            ftp_config = FTP_CONFIG['MovilsolAfuera']
        elif dd.value == "Digitar":
            ftp_config = FTP_CONFIG['Digitar']

        try:
            ftp = FTP()
            ftp.connect(ftp_config['host'])
            ftp.login(ftp_config['user'], ftp_config['password'])
            connection_status.value = f"Conexión establecida con: {dd.value}"
            print('Conexión Establecida')
            #print('Nos encontramos en la carpeta: ' + ftp.pwd())
            ftp.cwd(ftp_config['carpeta'])
            #print('Ahora nos encontramos en la carpeta: ' + ftp.pwd())
            #ftp.dir()

        except Exception as e:
            connection_status.value = f"No se pudo conectar: {str(e)}"
            print('No se pudo conectar: ' + str(e))
        
        page.update()

    t = ft.Text()
    dd = ft.Dropdown(
        on_change=dropdown_changed,
        options=[
            ft.dropdown.Option("Movilsol Oficina"),
            ft.dropdown.Option("Movilsol desde Afuera"),
            ft.dropdown.Option("Digitar"),
        ],
        width=200,
    )

    file_picker_save = ft.FilePicker(on_result=on_file_picker_result)
    page.overlay.append(dlg_modal_help)
    page.overlay.append(file_picker_save)

    page.add(
        ft.Column(
            [
                ft.Row(
                    [dd],
                    alignment=ft.MainAxisAlignment.CENTER,
                ),
                ft.Row(
                    [
                        path_var,
                        ft.ElevatedButton(text="Explorar", 
                                          icon="FOLDER_OUTLINED",
                                          width=150,
                                          on_click=lambda e: file_picker_save.save_file(file_type=ft.FilePickerFileType.CUSTOM, allowed_extensions=["xlsx"])
                                          ),
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                ),
                ft.Row(
                    [
                        ft.ElevatedButton(text="Aceptar", 
                                          icon="DONE_ALL",
                                          width=150,
                                          on_click=accept),
                        ft.ElevatedButton(text="Salir", 
                                          icon="EXIT_TO_APP",
                                          width=150,
                                          on_click=lambda e: page.window.close()
                                          ),
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                ),
                connection_status,
                result_label,
                data_table,
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        )
    )

ft.app(target=main)
