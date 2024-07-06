import os
import zipfile
import pandas as pd
from datetime import datetime
import flet as ft

#VAMOS A RECORRER LA CARPETA PARA LISTAR LA INFORMACION DENTRO DEL ARCHIVO .ZIP
def get_files_info(root_path, progress_callback):
    data = []
    for dirpath, _, filenames in os.walk(root_path):
        for filename in filenames:
            if filename.endswith('.zip'):
                zip_path = os.path.join(dirpath, filename)
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    for file_info in zip_ref.infolist():
                        file_date = datetime(*file_info.date_time).strftime('%Y-%m-%d %H:%M:%S')
                        data.append([os.path.basename(dirpath), file_info.filename, file_date])
        progress_callback(f"Verificando: {dirpath}")
    return data

#GUARDAMOS EL EXCEL EN EL DESTINO ELEGIDO
def save_to_excel(data, save_path):
    if not save_path.endswith('.xlsx'):
        save_path += '.xlsx'
    df = pd.DataFrame(data, columns=['Carpeta', 'Archivo en ZIP', 'Fecha'])
    df.to_excel(save_path, index=False, engine='openpyxl')
    return save_path

#CREAMOS LA VENTANA
def main(page: ft.Page):
    page.title = "Verificar Fechas Archivos Zip - by JoseLu - 2024"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.bgcolor = "#00bfff"
    page.window.width = 700
    page.window.height = 200
    page.window.maximizable = False

    path_var = ft.TextField(label="Ruta de las carpetas", width=400)

    def on_pick_result(e: ft.FilePickerResultEvent):
        if e.path:
            path_var.value = e.path
            page.update()

    def on_save_result(e: ft.FilePickerResultEvent):
        if e.path:
            verify_and_save(e.path)

    def verify_and_save(save_path):
        root_path = path_var.value
        if root_path:
            result_label.value = "Estamos verificando cada carpeta, aguarde un momento por favor."
            page.update()
            files_info = get_files_info(root_path, lambda msg: update_progress(msg))
            xlsx_path = save_to_excel(files_info, save_path)
            if xlsx_path:
                result_label.value = f"Información guardada en:\n{xlsx_path}"
            else:
                result_label.value = "No se ha seleccionado una ubicación para guardar."
            page.update()

    def update_progress(msg):
        result_label.value = msg
        page.update()

    def explore(e):
        file_picker.get_directory_path()

    def accept(e):
        file_picker_save.save_file(file_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    result_label = ft.Text(value="", color=ft.colors.WHITE)

    file_picker = ft.FilePicker(on_result=on_pick_result)
    file_picker_save = ft.FilePicker(on_result=on_save_result)
    page.overlay.extend([file_picker, file_picker_save])

    page.add(
        ft.Column(
            [
                ft.Row(
                    [
                        path_var,
                        ft.ElevatedButton(text="Explorar", on_click=explore),
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                ),
                ft.Row(
                    [
                        ft.ElevatedButton(text="Aceptar", on_click=accept),
                        ft.ElevatedButton(text="Salir", on_click=lambda e: page.window_close()),
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                ),
                result_label,
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        )
    )

ft.app(target=main)