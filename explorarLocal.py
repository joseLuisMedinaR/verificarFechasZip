import os
import zipfile
import pandas as pd
from datetime import datetime, timedelta
import flet as ft

# Función para obtener información de archivos dentro de archivos ZIP en una carpeta raíz dada
def get_files_info(root_path, progress_callback):
    data = []
    for dirpath, _, filenames in os.walk(root_path):
        for filename in filenames:
            if filename.endswith('.zip'):
                zip_path = os.path.join(dirpath, filename)
                zip_date = datetime.fromtimestamp(os.path.getmtime(zip_path)).strftime('%Y-%m-%d')
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    for file_info in zip_ref.infolist():
                        file_date = datetime(*file_info.date_time).strftime('%Y-%m-%d')
                        file_hour = datetime(*file_info.date_time).strftime('%H:%M:%S')
                        data.append([os.path.basename(dirpath), filename, zip_date, file_date, file_info.filename, file_hour])
        progress_callback(f"Verificando: {dirpath}")
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
    #page.window.full_screen = True
    page.window.maximizable = True
    page.window.minimizable = True
    page.scroll = True
    page.window_maximized = True

    path_var = ft.TextField(label="Ruta de las carpetas", width=400, disabled=True)
    result_label = ft.Text(value="", color=ft.colors.WHITE)

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
        #column_spacing=200,
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

    # Función que se ejecuta cuando se selecciona una carpeta
    def on_pick_result(e: ft.FilePickerResultEvent):
        if e.path:
            path_var.value = e.path
            page.update()

    # Función que se ejecuta cuando se selecciona una ubicación para guardar el archivo
    def on_save_result(e: ft.FilePickerResultEvent):
        if e.path:
            verify_and_save(e.path)

    # Función que verifica y guarda los datos en un archivo de Excel
    def verify_and_save(save_path):
        nonlocal xlsx_path  # Asegurar que xlsx_path sea modificable dentro de la función
        root_path = path_var.value
        if root_path:
            result_label.value = "Estamos verificando cada carpeta, aguarde un momento por favor."
            page.update()
            files_info = get_files_info(root_path, lambda msg: update_progress(msg))
            xlsx_path = save_to_excel(files_info, save_path)  # Asignar xlsx_path con la ruta de guardado
            if xlsx_path:
                result_label.value = f"Información guardada en:\n{xlsx_path}"
                nonlocal df
                df, data_table.rows = load_data_from_excel(xlsx_path)
                page.update()
            else:
                result_label.value = "No se ha seleccionado una ubicación para guardar."
            page.update()
        else:
            open_dlg_modal_help()

    # Función para actualizar el progreso
    def update_progress(msg):
        result_label.value = msg
        page.update()

    # Función para abrir el selector de carpetas
    def explore(e):
        file_picker.get_directory_path()

    # Función para abrir el selector de archivos para guardar
    def accept(e):
        if not path_var.value:
            open_dlg_modal_help()
        else:
            file_picker_save.save_file(file_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # Función para ordenar la tabla por una columna específica
    def sort_table(column_name, excel_path):
        nonlocal df
        ascending = not sort_state[column_name]  # Cambiar el estado de orden ascendente/descendente
        df, _ = load_data_from_excel(excel_path)  # Cargar de nuevo los datos desde el archivo Excel
        df = df.sort_values(by=column_name, ascending=ascending)
        df.to_excel(excel_path, index=False, engine='openpyxl')  # Sobrescribir el archivo Excel ordenado
        _, data_table.rows = load_data_from_excel(excel_path)  # Recargar las filas basadas en el DataFrame ordenado
        sort_state[column_name] = ascending  # Actualizar el estado de ordenamiento
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

    file_picker = ft.FilePicker(on_result=on_pick_result)
    file_picker_save = ft.FilePicker(on_result=on_save_result)
    page.overlay.extend([file_picker, file_picker_save])
    page.overlay.append(dlg_modal_help)

    page.add(
        ft.Column(
            [
                ft.Row(
                    [
                        path_var,
                        ft.ElevatedButton(text="Explorar", 
                                          icon="FOLDER_OUTLINED",
                                          width=150,
                                          on_click=explore),
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
                result_label,
                data_table,
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        )
    )

ft.app(target=main)
