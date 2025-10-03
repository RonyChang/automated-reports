from collections import Counter
from xlwings import Sheet, Range
from datetime import datetime

from classes.sub_section import SubSection

class Section:
    def __init__(self, name: str, data_sheet: Sheet, template_sheet: Sheet):
        self.name: str = name
        self.data_sheet: Sheet = data_sheet
        self.have_sub_sections: bool = None
        self.template_sheet: Sheet = template_sheet
        self.kpi_positions: list[tuple[str, int]] = None
        self.number_of_brands: int = None
        self.sub_section_positions: list[tuple[str, int]] = None
        self.sub_sections: list[SubSection] = []
        self.time_periods: list[str] = []
        self.last_column: int = None
        self.last_row: int = None

    def start_section(self, report_name: str):
        print(f"Starting section: {self.name}")
        self.get_positions_of_kpis()
        self.check_if_section_has_sub_sections()
        self.get_positions_of_sub_sections()
        self.get_time_periods()
        self.get_number_of_brands(report_name=report_name)
        self.create_sub_sections(report_name=report_name)
        self.get_sub_sections_rows()

    def get_positions_of_kpis(self):
        cells_row = self.data_sheet.range("1:1").value
        cells = [
            (cell_name, index + 1)
            for index, cell_name in enumerate(cells_row)
            if cell_name is not None
        ]
        self.kpi_positions = cells

    def get_time_periods(self):
        ''' Cambio 3'''
        # Lee toda la fila 2 de una sola vez (hasta la última columna usada) para evitar múltiples llamadas COM.
        last_col = self.data_sheet.range((2, self.data_sheet.cells.last_cell.column)).end('left').column
        values = self.data_sheet.range((2, 2), (2, last_col)).options(ndim=1).value
        # Normaliza fechas (datetime -> 'Mes-AñoCorto') y mantiene strings como están
        self.time_periods = []
        for v in values:
            if v is None:
                break
            if isinstance(v, datetime):
                v = v.strftime("%b-%y")
            self.time_periods.append(v)
        print("Periodos normalizados...")


    def check_if_section_has_sub_sections(self):
        ''' Cambio 4'''
        # Evita leer toda la columna A:A (1M+ filas). Lee solo hasta la última fila usada.
        last_row = self.data_sheet.range((self.data_sheet.cells.last_cell.row, 1)).end('up').row
        cells_column = self.data_sheet.range((1, 1), (last_row, 1)).options(ndim=1).value
        cells = [
            (str(cell_name).strip(), index + 1)
            for index, cell_name in enumerate(cells_column)
            if cell_name is not None
        ]
        positions = []
        for cell in cells:
            next_cell = self.data_sheet.range((cell[1], 2)).value
            if next_cell is None and cell[1] != 1:
                positions.append(cell)
        self.have_sub_sections = len(positions) > 1
        #self.have_sub_sections = True if self.name != "marcas" else False
        if self.have_sub_sections:
            print(f"\t-Section {self.name} WITH sub sections")
        else:
            print(f"\t-Section {self.name} WITHOUT sub sections")


    def get_positions_of_sub_sections(self):
        ''' Cambio 5'''
        if not self.have_sub_sections:
            return
        last_row = self.data_sheet.range((self.data_sheet.cells.last_cell.row, 1)).end('up').row
        cells_column = self.data_sheet.range((1, 1), (last_row, 1)).options(ndim=1).value
        cells = [
            (str(cell_name).strip(), index + 1)
            for index, cell_name in enumerate(cells_column)
            if cell_name is not None
        ]
        positions = []
        for cell in cells:
            next_cell = self.data_sheet.range((cell[1], 2)).value
            if next_cell is None and cell[1] != 1:
                positions.append(cell)
        self.sub_section_positions = positions

    def get_number_of_brands(self, report_name: str):
        """
        Calcula number_of_brands (para la 1ª subsección) y además
        brands_per_subsection = [N_0, N_1, ...] contando desde la fila
        debajo del header hasta la primera A vacía (sin pasar el siguiente header).
        """
        def is_blank(v):
            return v is None or (isinstance(v, str) and v.strip() == "")

        # última fila usada en A
        last_row = self.data_sheet.range((self.data_sheet.cells.last_cell.row, 1)).end('up').row
        colA = self.data_sheet.range((1, 1), (last_row, 1)).options(ndim=1).value
        if not isinstance(colA, list):
            colA = [colA]

        # headers de subsecciones (fila del título en A)
        if self.have_sub_sections and self.sub_section_positions:
            starts = [pos[1] for pos in self.sub_section_positions]
        else:
            starts = [3]  # header único en A3

        counts = []
        for i, s in enumerate(starts):
            start = s + 1
            end_guard = (starts[i+1] - 1) if i + 1 < len(starts) else last_row
            r, c = start, 0
            while r <= end_guard:
                a = colA[r-1] if r-1 < len(colA) else None  # 1-based -> 0-based
                if is_blank(a):
                    break
                c += 1
                r += 1
            counts.append(c)

        self.brands_per_subsection = counts
        self.number_of_brands = counts[0] if counts else 0
        print(f"\t- brands_per_subsection={self.brands_per_subsection} | number_of_brands={self.number_of_brands}")

    def create_sub_sections(self, report_name: str):
        if not self.have_sub_sections:
            sub_section_name = f"Total {self.name.capitalize()}"
            sub_section = SubSection(name=sub_section_name, starting_row=2)
            self.sub_sections.append(sub_section)
            return
        for sub_section_position in self.sub_section_positions:
            name, starting_row = sub_section_position
            print(f"Creating sub-section - {name}")
            sub_section = SubSection(name=name, starting_row=starting_row)
            self.sub_sections.append(sub_section)

    def get_sub_sections_rows(self):
        """
        Asigna target_row con alturas variables por subsección.
        Bloque por subsección en plantilla = 1 header + N marcas + 2 separadores (= 3 + N).
        """
        brands_list = getattr(self, "brands_per_subsection", None)
        if not brands_list:
            brands_list = [self.number_of_brands] * len(self.sub_sections)

        row = 5  # la 1ª subsección empieza en fila 5 de la plantilla
        for i, sub in enumerate(self.sub_sections):
            sub.target_row = row
            n = brands_list[i] if i < len(brands_list) else self.number_of_brands
            row += 3 + n

    def insert_columns_for_time_periods(self):
        start_column = 3
        last_row = self.template_sheet.cells.last_cell.row
        for index, period in enumerate(self.time_periods):
            if index != len(self.time_periods) - 1:
                prev_column = self.template_sheet.range(
                    (1, start_column), (last_row, start_column)
                )
                prev_column.insert(shift="right")
            self.template_sheet.range((3, start_column - 1)).value = period
            self.template_sheet.range(
                (3, start_column - 1), (4, start_column - 1)
            ).merge()
            start_column += 1

    def create_kpis_sections(self):
        print("Creating sections for kpi's")
        kpis_list = self.sub_sections[0].kpis
        for index, kpi in enumerate(kpis_list):
            print(f"\t-Creating section for kpi: {kpi.report_name}")
            self.copy_base_kpi_section(
                kpi_name=kpi.report_name,
                template_sheet=self.template_sheet,
                index=index + 1,
            )
        self.delete_base_kpi_section()

    def copy_base_kpi_section(self, kpi_name: str, template_sheet: Sheet, index: int):
        """
        Copia las columnas de periodos, variación y nombre de kpi por cada kpi del reporte
        """
        last_row = template_sheet.cells.last_cell.row
        time_periods_number = len(self.time_periods)
        start_column = 2

        last_column = (
            start_column + time_periods_number - 1 + 3 + 1
        )  # Hace unos minutos sabía el significado de estos números, ahora no

        base_section = template_sheet.range((1, start_column), (last_row, last_column))

        # Formula para obtener la posición del siguiente kpi. Culpa de Estef
        increment = time_periods_number + 4
        kpi_section_start_column = index * increment + 2

        kpi_section_end_column = kpi_section_start_column + time_periods_number + 3
        kpi_section = template_sheet.range((1, kpi_section_start_column))
        base_section.copy()
        kpi_section.paste()
        template_sheet.api.Application.CutCopyMode = False
        template_sheet.range((1, kpi_section_end_column)).value = kpi_name

    def delete_last_column(self):
        """
        Por algún motivo al copiar y pegar las secciones de los kpis,
        la última sección copia una columna adicional.
        Esta función elimina esa columna. Perdón pero es más fácil hacer esto que arreglar el bug
        Actualización 22/04/2025:
            Esta función ya no es necesaria, pero se deja como recuerdo.
        """
        row = self.template_sheet.range("3:3").value
        cells = [
            (index, cell_name)
            for index, cell_name in enumerate(row)
            if cell_name is not None
        ]
        last_column = cells[-1][0] + 1
        self.template_sheet.range((1, last_column)).api.EntireColumn.Delete()

    def delete_base_kpi_section(self):
        """
        Elimina el template de kpi
        """
        start_column = 2
        last_column = 2 + len(self.time_periods) + 3
        self.template_sheet.range(
            (1, start_column), (1, last_column)
        ).api.EntireColumn.Delete()

    def get_data(self):
        for i, sub_section in enumerate(self.sub_sections):
            # usa el conteo específico de esa subsección
            n = (
                self.brands_per_subsection[i]
                if hasattr(self, "brands_per_subsection")
                and i < len(self.brands_per_subsection)
                else self.number_of_brands
            )
            print(f"\t\t-Getting data for sub section: {sub_section.name} (brands={n})")
            sub_section.get_data(
                sheet=self.data_sheet,
                kpi_positions=self.kpi_positions,
                time_periods=self.time_periods,
                number_of_brands=n,
            )


    def create_sub_sections_on_sheet(self):
        for sub_section in self.sub_sections:
            sub_section.create_on_sheet(
                sheet=self.template_sheet,
            )

    def get_kpi_positions(self):
        for sub_section in self.sub_sections:
            sub_section.get_kpi_positions(sheet=self.template_sheet)

    def add_var_columns_to_df(
        self, first_difference: tuple[str, str], second_difference: tuple[str, str], third_difference: tuple[str, str]
    ):
        for sub_section in self.sub_sections:
            for kpi in sub_section.kpis:
                kpi.add_var_columns_to_df(
                    first_difference=first_difference,
                    second_difference=second_difference,
                    third_difference=third_difference,
                )

    def copy_data(self):
        time_periods_number = len(self.time_periods)
        for sub_section in self.sub_sections:
            print(f"\t\t-Copying data for sub section: {sub_section.name}")
            sub_section.copy_data(
                sheet=self.template_sheet,
                time_periods_number=time_periods_number,
            )

    def copy_brands(self):
        for sub_section in self.sub_sections:
            print(f"\t\t-Copying brands for sub section: {sub_section.name}")
            sub_section.copy_brands(sheet=self.template_sheet)

    def get_last_column(self):
        cells_row = self.template_sheet.range("1:1").value
        cells = [
            (cell_name, index + 1)
            for index, cell_name in enumerate(cells_row)
            if cell_name is not None
        ]
        self.last_column = cells[-1][1]

    def get_last_row(self):
        cells_column = self.template_sheet.range("A:A").value
        cells = [
            (cell_name.strip(), index + 1)
            for index, cell_name in enumerate(cells_column)
            if cell_name is not None
        ]
        self.last_row = cells[-1][1]

    def style_report(self):
        time_periods_number = len(self.time_periods)
        print(f"\t-Styling section: {self.name}")
        self.get_last_row()
        self.get_last_column()
        self.style_variation_columns(time_periods_number=time_periods_number)
        self.style_segments_rows()
        self.style_kpi_columns(time_periods_number=time_periods_number)

    def style_segments_rows(self):
        """Cambio 7"""
        # Encuentra la última fila con datos en la columna A
        last_row = self.template_sheet.range(
            "A" + str(self.template_sheet.cells.last_cell.row)
        ).end("up").row
        # Lee solo hasta la última fila con datos
        cells_column = self.template_sheet.range(f"A1:A{last_row}").value
        # Aplanar (xlwings devuelve lista de listas en columnas)
        if isinstance(cells_column[0], list):
            cells_column = [c[0] if c else None for c in cells_column]
        # Iterar por filas y detectar las que empiezan con "*"
        for row_number, cell_text in enumerate(cells_column, start=1):
            if cell_text and str(cell_text).strip().startswith("*"):
                self.paint_segment_row(row_number=row_number)
                self.change_segment_cell_text(row_number=row_number)

    def paint_segment_row(self, row_number: int):
        """
        Pinta la fila de una sección con un azul claro
        """
        row_color = (0, 181, 255)  ##82c8a6
        row = self.template_sheet.range((row_number, 1), (row_number, self.last_column))
        row.color = row_color
        row.font.color = (255, 255, 255)

    def change_segment_cell_text(self, row_number: int):
        """
        Quita el asterisco inicial del nombre de la sección
        """
        cell = self.template_sheet.range((row_number, 1))
        cell.value = cell.value.replace("*", "")

    def style_kpi_columns(self, time_periods_number: int):
        """
        Cambia el color de fondo de las columnas de los kpis a gris
        """
        section_kpis = self.sub_sections[0].kpis
        for kpi in section_kpis:
            kpi_column = kpi.column_pos
            if kpi_column is None:
                continue
            kpi_column = self.template_sheet.range(
                (6, kpi_column), (self.last_row, kpi_column)
            )
            self.paint_kpi_column(
                column=kpi_column, time_periods_number=time_periods_number
            )

    def paint_kpi_column(self, column: Range, time_periods_number: int):
        base_cell = self.template_sheet.range(5, 4 + time_periods_number)
        base_cell.copy(column)

    def style_variation_columns(self, time_periods_number: int):
        """
        Agrega el formato de variación (color de fondo y color de texto)
        """
        section_kpis = self.sub_sections[0].kpis
        for kpi in section_kpis:
            kpi_column = kpi.column_pos
            if kpi_column is None:
                continue
            var_range = self.template_sheet.range(
                (7, kpi_column - 3), (self.last_row, kpi_column - 1)
            )
            self.paint_var_columns(
                time_periods_number=time_periods_number,
                var_range=var_range,
            )

    def paint_var_columns(
        self,
        time_periods_number: int,
        var_range: Range,
    ):
        base_var_cell = self.template_sheet.range(6, time_periods_number + 2)

        base_var_cell.copy()
        var_range.paste(paste="formats")
