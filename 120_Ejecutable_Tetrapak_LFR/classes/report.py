import os
import shutil
from datetime import datetime
from xlwings import Book
import xlwings as xw

from .section import Section

class Report:
    def __init__(
        self,
        name: str,
        period: str,
        first_difference: tuple,
        second_difference: tuple,
        third_difference: tuple,
    ):
        self.name: str = name
        self.period: str = period
        self.current_year: int = None
        self.sections: list[Section] = []
        self.data_book: Book = Book("data.xlsx")
        self.template_book: Book = None
        self.first_difference: tuple = first_difference
        self.second_difference: tuple = second_difference
        self.third_difference: tuple = third_difference
        self.last_column: int = None

        self.init_report()

    def init_report(self):
        self.get_current_year()
        self.copy_template_book()
        self.open_template_book()
        self.change_var_text()
        self.insert_columns_for_time_periods()
        self.create_kpis_sections()

    def copy_template_book(self):
        report_name = f"{self.name} {self.period}-{self.current_year}.xlsx"
        if os.path.exists(report_name):
            os.remove(report_name)
        shutil.copy('template.xlsx', report_name)
    
    def open_template_book(self):
        report_name = f"{self.name} {self.period}-{self.current_year}.xlsx"
        self.template_book: Book = Book(report_name)

    def get_current_year(self):
        """ "
        Esto va dar el año incorrento cuando sea la liberación del Q4,
        quien mantenga este código está en la autoridad moral de corregirlo
        """
        """
        Ya lo corregí antes de que sea Q4 😇
        El futuro soy yo ----> By Vir
        """
        if datetime.now().month == 1:
            self.current_year = datetime.now().year - 1
        else:
            self.current_year = datetime.now().year

    def insert_columns_for_time_periods(self):
        for section in self.sections:
            section.insert_columns_for_time_periods()

    def create_kpis_sections(self):
        for section in self.sections:
            section.create_kpis_sections()

    def change_var_text(self):
        sheet = self.template_book.sheets["template"]
        first_var_cell = sheet.range("C4")
        second_var_cell = sheet.range("D4")
        third_var_cell = sheet.range("E4")

        first_var_cell.value = (
            f"{self.first_difference[0]} vs {self.first_difference[1]}"
        )
        first_var_cell.api.WrapText = True

        second_var_cell.value = (
            f"{self.second_difference[0]} vs {self.second_difference[1]}"
        )
        second_var_cell.api.WrapText = True
        third_var_cell.value = (
            f"{self.third_difference[0]} vs {self.third_difference[1]}"
        )
        third_var_cell.api.WrapText = True

    def delete_base_kpi_section(self):
        """
        Elimina el template de kpi
        """
        template_sheet = self.template_book.sheets["template"]
        template_sheet.delete()

    def create_sections(self):
        report_sections = self.get_report_sections()
        for section_name in report_sections:
            section_sheet = self.data_book.sheets[section_name]
            template_sheet = self.copy_template_sheet_for_section(
                section_name=section_name
            )
            section = Section(
                name=section_name,
                data_sheet=section_sheet,
                template_sheet=template_sheet,
            )
            self.sections.append(section)

    def copy_template_sheet_for_section(self, section_name: str):
        template_sheet = self.template_book.sheets["template"]
        template_sheet.copy(name=section_name)
        return self.template_book.sheets[section_name]

    def get_report_sections(self):
        return self.data_book.sheet_names

    def start_sections(self):
        for section in self.sections:
            section.start_section(report_name=self.name)

    def get_data(self):
        print("Getting data")
        for section in self.sections:
            print(f"\t-Getting data for section: {section.name}")
            section.get_data()

    def create_sub_sections_on_sheet(self):
        for section in self.sections:
            section.create_sub_sections_on_sheet()

    def get_kpi_positions(self):
        for section in self.sections:
            section.get_kpi_positions()

    def add_var_columns_to_df(self):
        for section in self.sections:
            section.add_var_columns_to_df(
                first_difference=self.first_difference,
                second_difference=self.second_difference,
                third_difference=self.third_difference,
            )

    def copy_data(self):
        print("Copying data")
        for section in self.sections:
            print(f"\t-Copying data for section: {section.name}")
            section.copy_data()

    def copy_brands(self):
        for section in self.sections:
            print(f"\t-Copying brands for section: {section.name}")
            section.copy_brands()

    def style_report(self):
        print("Styling report")
        self.write_category_name_on_index()
        self.write_period_on_index()
        self.remove_template_sheet()
        for section in self.sections:
            section.style_report()

    def write_category_name_on_index(self):
        index_sheet = self.template_book.sheets["Indice"]
        index_sheet.range("F7").value = self.name

    def write_period_on_index(self):
        index_sheet = self.template_book.sheets["Indice"]
        index_sheet.range("J9").value = f"{self.period}-{self.current_year}"

    def remove_template_sheet(self):
        self.template_book.sheets["template"].delete()

    def add_hyperlinks(self):
        for index, section in enumerate(self.sections):
            self.add_hyperlink_to_section(section=section, index=index)

    def add_hyperlink_to_section(self, section: Section, index: int):
        row = 15 + (index * 2)
        index_sheet = self.template_book.sheets["Indice"]

        # Copiar fila base
        if index != 0:
            base_row = index_sheet.range("15:15")
            output_row = index_sheet.range(f"{row}:{row}")
            base_row.copy(output_row)

        # Rango objetivo (H:J en la fila)
        target_rng = index_sheet.range(f"H{row}:J{row}")
        try:
            # Si estuviera combinado, descombinar primero
            target_rng.api.UnMerge()
        except Exception:
            pass

        # Texto a mostrar en H{row}
        cell_h = index_sheet.range(f"H{row}")
        cell_h.value = section.name
        safe_name = section.name.replace("'", "''")  # manejar nombres con '

        # Hiperlink en H{row}
        index_sheet.api.Hyperlinks.Add(
            Anchor=cell_h.api,
            Address="",
            SubAddress=f"'{safe_name}'!A1",
            ScreenTip="",
            TextToDisplay=section.name,
        )

        # Combinar H{row}:J{row}
        try:
            target_rng.api.Merge()
        except Exception as e:
            print(f"[WARN] No se pudo combinar H{row}:J{row}: {e}")

        # Formato: blanco + negrita sobre el área combinada
        try:
            target_rng.api.Font.Bold = True
            target_rng.api.Font.Color = 16777215  # blanco
            # xlCenter = -4108
            target_rng.api.HorizontalAlignment = -4108
            target_rng.api.VerticalAlignment = -4108
        except Exception as e:
            print(f"[WARN] No se pudo aplicar formato en H{row}:J{row}: {e}")
    

    def clean_subsections(self):
        """
        Agregado 1
        Elimina desde la primera fila en blanco hasta una fila antes del inicio de la siguiente subsección,
        y además elimina desde la primera celda en blanco debajo de la última subsección hasta el final.
        El borrado de la última subsección se hace primero para evitar corrimientos.
        """
        print("Limpiando subsecciones...")

        for section in self.sections:
            sheet = section.template_sheet
            subs = section.sub_sections
            if not subs:
                continue

            print(f"\t• Hoja: {sheet.name} — {len(subs)} subsecciones")

            # Manejo especial para la última subsección (se hace PRIMERO)
            last_sub = subs[-1]
            last_start_row = last_sub.target_row
            #last_row_sheet = sheet.cells.last_cell.row  # fin absoluto de la hoja
            last_row_sheet = sheet.api.UsedRange.Row + sheet.api.UsedRange.Rows.Count - 1 # última fila usada

            # Se busca desde la fila siguiente a la subsección hasta el final de la hoja
            rng = sheet.range((last_start_row+1, 1), (last_row_sheet, 1)).options(ndim=1).value
            if not isinstance(rng, list):
                rng = [rng]

            first_empty = None
            for offset, val in enumerate(rng, start=last_start_row+1):
                if val is None:
                    first_empty = offset
                    break

            if first_empty is not None:
                delete_from = first_empty
                delete_to = last_row_sheet
                print(f"\t\t- Última subsección '{last_sub.name}': eliminando {delete_from}:{delete_to}")
                sheet.api.Rows(f"{delete_from}:{delete_to}").Delete()
            else:
                print(f"\t\t- Última subsección '{last_sub.name}': sin filas vacías debajo.")

            # Limpiamos subsecciones intermedias
            delete_ranges = []

            for i, sub in enumerate(subs):
                start_row = sub.target_row
                next_row = subs[i+1].target_row if i + 1 < len(subs) else None

                search_start = start_row + 1
                search_end = (next_row - 1) if next_row else sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row

                if search_start > search_end:
                    continue

                rng = sheet.range((search_start, 1), (search_end, 1)).options(ndim=1).value
                if not isinstance(rng, list):
                    rng = [rng]

                first_empty = None
                for offset, val in enumerate(rng, start=search_start):
                    if val is None:
                        first_empty = offset
                        break

                if first_empty is not None:
                    delete_from = first_empty
                    delete_to = search_end
                    if delete_from <= delete_to:
                        delete_ranges.append((delete_from, delete_to))
                        print(f"\t\t- '{sub.name}': marcado para eliminar {delete_from}:{delete_to-1}")
                else:
                    print(f"\t\t- '{sub.name}': sin celdas vacías en rango.")

            # Borrar rangos desde abajo hacia arriba
            for delete_from, delete_to in reversed(delete_ranges):
                print(f"\t\t Eliminando filas {delete_from}:{delete_to-1}")
                sheet.api.Rows(f"{delete_from}:{delete_to-1}").Delete()

        print("Limpieza finalizada.")

    def final_styling(self):
        """
        Reestiliza las segundas filas de cada subsección:
        - Fondo celeste (0, 181, 255)
        - Texto en negrita
        - Texto en color blanco
        """
        print("Aplicando estilo a subsecciones...")

        row_color = (0, 230, 186)  
        font_white = 16777215        # blanco

        for section in self.sections:
            sheet = section.template_sheet
            subs = section.sub_sections
            if not subs:
                continue

            print(f"\t• Hoja: {sheet.name} — {len(subs)} subsecciones")

            for sub in subs:
                second_row = sub.target_row + 1

                # Calcular la última columna usada en la hoja
                last_col = sheet.api.UsedRange.Columns.Count

                try:
                    rng = sheet.range((second_row, 1), (second_row, last_col))
                    rng.color = row_color
                    rng.api.Font.Bold = True
                    rng.api.Font.Color = font_white
                    print(f"\t\t- Estilada fila {second_row} hasta col {last_col} en '{sub.name}'")
                except Exception as e:
                    print(f"\t\t[ERROR] No se pudo estilizar fila {second_row} en '{sub.name}': {e}")

        print("Estilado finalizado.")
        
    def remove_first_rows(self):
        """
        En cada hoja:
        1) Borra SIEMPRE la fila 5.
        2) Desde la fila 5 hacia abajo, busca patrones:
            [fila vacía en A] -> [siguiente fila con texto en A]  => ELIMINAR esa siguiente fila (header).
            Además, se BORRA EL FORMATO de la fila vacía separadora completa.
            Se detiene cuando encuentra dos filas vacías consecutivas (para no recorrer toda la hoja).

        Dos pasadas: primero marco (y limpio formato de) separadores, luego borro en orden inverso.
        """
        print("Eliminando primeras filas de subsecciones...")

        def a_is_empty(sheet, row):
            v = sheet.range((row, 1)).value
            if v is None:
                return True
            if isinstance(v, str):
                return v.strip() == ""
            return False  # otros tipos (número/fecha) => no vacío

        # Hojas únicas desde las sections
        sheets = {}
        for section in getattr(self, "sections", []):
            sh = section.template_sheet
            sheets.setdefault(sh.name, sh)

        for name, sheet in sheets.items():
            print(f"\t• Hoja: {name}")

            # Última fila usada aproximada (para acotar el escaneo)
            try:
                used = sheet.api.UsedRange
                last_row = used.Row + used.Rows.Count - 1
            except Exception:
                last_row = sheet.range((sheet.cells.last_cell.row, 1)).end("up").row

            rows_to_delete = set()
            separator_rows_to_clear = set()

            # 1) Siempre eliminar fila 5
            rows_to_delete.add(5)

            # 2) Escaneo: detectar [vacío] -> [texto] y
            #    - marcar esa fila "texto" para borrar
            #    - limpiar formato de la fila vacía separadora
            r = 5
            while r <= last_row:
                # si dos vacías seguidas: detener
                if a_is_empty(sheet, r) and a_is_empty(sheet, r + 1 if r + 1 <= last_row else last_row):
                    print(f"\t\t- Doble vacío en A{r}:A{r+1}. Deteniendo escaneo.")
                    break

                if a_is_empty(sheet, r):
                    nxt = r + 1
                    if nxt <= last_row and not a_is_empty(sheet, nxt):
                        # Vacío -> Texto: limpiar formato de la fila vacía (separador)
                        separator_rows_to_clear.add(r)
                        # y borrar el header (nxt)
                        rows_to_delete.add(nxt)
                        print(f"\t\t- Marcada para eliminar fila {nxt} (header tras vacío A{r})")
                        r = nxt + 1
                        continue

                r += 1

            # 2.5) Borrar FORMATO de todas las filas separadoras detectadas (antes de borrar filas)
            for row in sorted(separator_rows_to_clear):
                try:
                    sheet.api.Rows(f"{row}:{row}").ClearFormats()
                    print(f"\t\t- Formato borrado en fila {row} (separador).")
                except Exception as e:
                    print(f"\t\t[WARN] No se pudo borrar formato en fila {row}: {e}")

            # 3) Borrar filas marcadas en orden inverso (evita corrimientos)
            if rows_to_delete:
                for row in sorted(rows_to_delete, reverse=True):
                    try:
                        sheet.api.Rows(f"{row}:{row}").Delete()
                        print(f"\t\t- Eliminada fila {row}")
                    except Exception as e:
                        print(f"\t\t[WARN] No se pudo eliminar fila {row}: {e}")

        print("Eliminación de primeras filas completada.")


    def apply_star_bold_underline(self, plantilla_path="LFR Plantilla.xlsx"):
        """
        Post-procesado final:
        Para cada hoja que exista en el reporte y en 'LFR Plantilla.xlsx',
        si en la plantilla la columna C (fila r) contiene '*', entonces
        en el reporte se formatea la fila (r + 2) completa con negrita y subrayado.
        """

        print("Aplicando formato (negrita + subrayado) según '*' en plantilla...")

        # --- abrir / obtener la plantilla ---
        plantilla_book = None
        try:
            # intenta reutilizar el libro si ya está abierto
            for app in xw.apps:
                for bk in app.books:
                    if bk.name.lower() == os.path.basename(plantilla_path).lower():
                        plantilla_book = bk
                        break
                if plantilla_book:
                    break
            if plantilla_book is None:
                plantilla_book = xw.Book(plantilla_path)
        except Exception as e:
            print(f"[ERROR] No se pudo abrir '{plantilla_path}': {e}")
            return

        # --- mapear hojas destino del reporte ---
        dest_sheets = {}
        for section in getattr(self, "sections", []):
            sh = section.template_sheet
            dest_sheets[sh.name] = sh  # evita duplicados por nombre

        plantilla_sheet_names = {sh.name: sh for sh in plantilla_book.sheets}

        # Constante Excel para subrayado simple
        XL_UNDERLINE_SINGLE = 2

        for sh_name, dest_sh in dest_sheets.items():
            src_sh = plantilla_sheet_names.get(sh_name)
            if src_sh is None:
                print(f"\t• Hoja '{sh_name}': no existe en la plantilla; se omite.")
                continue

            print(f"\t• Hoja '{sh_name}': escaneando columna C en plantilla...")

            # última fila con datos en la columna C de la PLANTILLA
            try:
                # xlUp = -4162
                last_row_src = src_sh.api.Cells(src_sh.api.Rows.Count, 2).End(-4162).Row
            except Exception:
                last_row_src = src_sh.range((src_sh.cells.last_cell.row, 2)).end('up').row

            # última columna usada en el REPORTE, para formatear toda la fila
            try:
                used = dest_sh.api.UsedRange
                first_col = used.Column
                last_col_dest = first_col + used.Columns.Count - 1
            except Exception:
                last_col_dest = dest_sh.cells.last_cell.column

            # recolecta filas a formatear (evita borrar/editar mientras iteras)
            rows_to_format = []
            for r in range(1, max(1, last_row_src) + 1):
                val = src_sh.range((r, 2)).value
                if isinstance(val, str) and ("*" in val):
                    rows_to_format.append(r + 2)  # offset +2 hacia el reporte

            # aplicar formato en bloque (de una en una por robustez)
            for dr in rows_to_format:
                try:
                    rng = dest_sh.range((dr, 1), (dr, last_col_dest))
                    rng.api.Font.Bold = True
                    rng.api.Font.Underline = XL_UNDERLINE_SINGLE
                    # Si quieres también color del texto o fondo, puedes añadirlo aquí
                    print(f"\t\t- Fila {dr} formateada (marca en plantilla C{dr-2}).")
                except Exception as e:
                    print(f"\t\t[WARN] No se pudo formatear fila {dr}: {e}")

        print("Formato aplicado según plantilla.")

    def format_two_decimals_and_zero_dash(self):
        """
        Formato final del reporte:
        - Máximo 2 decimales (sin forzar .00).
        - Muestra '-' para 0 (incluye %).
        - No modifica datos.
        - No toca la columna A ni columnas con formatos de fecha/hora.
        """
        print("Aplicando formato: máximo 1 decimales + 0 como '-'...")

        # Reunir hojas únicas del reporte
        sheets = {}
        for section in getattr(self, "sections", []):
            sh = section.template_sheet
            sheets[sh.name] = sh

        for name, sheet in sheets.items():
            print(f"\t• Hoja: {name}")
            # Límites de rango usado
            try:
                used = sheet.api.UsedRange
                first_row = used.Row
                last_row  = used.Row + used.Rows.Count - 1
                first_col = used.Column
                last_col  = first_col + used.Columns.Count - 1
            except Exception:
                first_row = 1
                last_row  = sheet.range((sheet.cells.last_cell.row, 1)).end("up").row
                last_col  = sheet.cells.last_cell.column

            if last_col < 2 or last_row < first_row:
                print("\t\t- Sin rango aplicable.")
                continue

            data_start_row = max(first_row, 5)  # evita encabezados altos

            for col in range(2, last_col + 1):  # B..última
                # Buscar una muestra para inspeccionar el NumberFormat real de la columna
                sample_r = None
                base_nf = None
                for r in range(data_start_row, min(data_start_row + 300, last_row) + 1):
                    try:
                        base_nf = str(sheet.range((r, col)).api.NumberFormat)
                    except Exception:
                        base_nf = None
                    v = sheet.range((r, col)).value
                    # tomamos la primera celda no vacía, aunque sea texto; solo queremos el formato
                    if base_nf:
                        sample_r = r
                        break

                if not base_nf:
                    continue  # no hay nada útil en la columna

                nf_lower = base_nf.lower()

                # Evitar columnas de fecha/hora para no romper calendarios
                if any(tok in nf_lower for tok in ["yy", "dddd", "mmm", "dd", "mm", "hh", "ss", "am/pm"]):
                    # formatos típicos de fecha/hora: dejamos igual
                    continue

                # ¿Es porcentaje?
                is_percent = "%" in base_nf

                # Formatos con hasta 2 decimales y cero como '-'
                if is_percent:
                    new_nf = '0.0%;-0.0%;"-";@'
                else:
                    new_nf = '#,##0.0;-#,##0.0;"-";@'

                # Si ya está igual, no re-asignar
                if str(base_nf).strip() == new_nf.strip():
                    continue

                try:
                    col_rng = sheet.range((first_row, col), (last_row, col))
                    col_rng.api.NumberFormat = new_nf
                except Exception as e:
                    print(f"\t\t[WARN] No se pudo aplicar formato en col {col}: {e}")

        print("Formato aplicado: máx 2 decimales y ceros como '-'.")

    def paint_only_kpi_columns_post(self, color=(0, 83, 92)):
        """
        Post: pinta solo las columnas KPI con #00535C,
        desde la fila 5 hasta la última fila con datos en la columna A.
        """
        for section in getattr(self, "sections", []):
            sheet = section.template_sheet

            # última fila con datos en la columna A (xlUp = -4162)
            try:
                last_data_row = sheet.api.Cells(sheet.api.Rows.Count, 1).End(-4162).Row
            except Exception:
                # fallback si no hay API: usa xlwings .end('up')
                last_data_row = sheet.range((sheet.cells.last_cell.row, 1)).end("up").row

            start_row = 5
            if last_data_row < start_row:
                continue

            # KPIs de la primera subsección, igual que tu flujo
            if not section.sub_sections:
                continue
            kpis = getattr(section.sub_sections[0], "kpis", [])
            if not kpis:
                continue

            for kpi in kpis:
                col_idx = getattr(kpi, "column_pos", None)
                if not col_idx:
                    continue
                try:
                    sheet.range((start_row, col_idx), (last_data_row, col_idx)).color = color  # #00535C
                except Exception as e:
                    print(f"[WARN] {sheet.name}: no se pudo pintar col {col_idx} hasta fila {last_data_row}: {e}")

    def save(self):
        report_name = f"{self.name} {self.period}-{self.current_year}.xlsx"
        print(f"Saving report as {report_name}")
        self.template_book.save()