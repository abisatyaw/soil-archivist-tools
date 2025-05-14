import pandas as pd
import flet as ft
import engine

class MainApplication:
    def __init__(self):
        self.selected_file_path = None

    def run(self, page: ft.Page):
        page.title = "Excel Column Finder"
        page.window_width = 600
        page.window_height = 500
        page.scroll = ft.ScrollMode.AUTO

        file_picker = ft.FilePicker(on_result=self.on_file_pick)

        self.file_path_input = ft.TextField(
            label="Selected Excel File",
            read_only=True,
            width=500
        )

        self.project_name_input = ft.TextField(label="Project Title", width=300)
        self.generate_button = ft.ElevatedButton(text="Generate", on_click=self.on_generate_click)

        self.status_text = ft.Text(value="No file selected", color="grey")
        self.output_column_data = ft.TextField(
            label="Column Data",
            value=" ",
            multiline=True, 
            read_only=True,  
            min_lines=10,     
            max_lines=15,     
            expand=False,     
        )

        pick_button = ft.ElevatedButton(
            text="Choose Excel File",
            icon=ft.Icons.UPLOAD_FILE,
            on_click=lambda e: file_picker.pick_files(allow_multiple=False, allowed_extensions=["xlsx"])
        )

        page.overlay.append(file_picker)
        page.add(
            ft.Column(
                [
                    ft.Row([
                        pick_button,
                        self.file_path_input
                    ]),
                    ft.Row([
                        self.project_name_input]),
                    self.generate_button,
                    self.status_text,
                    self.output_column_data
                ],
                spacing=20,
                expand=True
            )
        )

    def on_file_pick(self, e: ft.FilePickerResultEvent):
        if e.files:
            self.selected_file_path = e.files[0].path
            self.file_path_input.value = e.files[0].name  # Show only filename
            self.status_text.value = "File selected. Enter sheet and column name to search."
            self.status_text.color = "blue"
        else:
            self.status_text.value = "No file selected"
            self.status_text.color = "red"
        self.file_path_input.update()
        self.status_text.update()

    def on_generate_click(self, e):
        self.status_text.value = "Please enter both sheet name and column name."

        engine.test()

if __name__ == "__main__":
    app = MainApplication()
    ft.app(target=app.run)