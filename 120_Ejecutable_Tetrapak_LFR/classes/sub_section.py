import pandas as pd
from xlwings import Sheet

from classes.kpi import Kpi

class SubSection:
    def __init__(self, name: str, starting_row: int):
        self.name: str = name
        self.starting_row: int = starting_row
        self.target_row: int = None
        self.kpis: list[Kpi] = []

    def __str__(self):
        return f"SubSection: {self.name} - Starting row: {self.starting_row}"
    
    """Código agregado""" # Función: que Python no lea como números si las fechas están en meses
    nombres_meses = {
        1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr",
        5: "May", 6: "Jun", 7: "Jul", 8: "Ago",
        9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"
    }

    def reformatear_fecha(self, fecha):
        try:
            fecha_parseada = pd.to_datetime(fecha)
            return f"{self.nombres_meses[fecha_parseada.month]}-{str(fecha_parseada.year)[-2:]}"
        except:
            return fecha
    """"""
    def get_data(
        self,
        sheet: Sheet,
        kpi_positions: list[tuple[str, int]],
        time_periods: list[str],
        number_of_brands: int,
    ):
        for kpi_data in kpi_positions:
            kpi_name, kpi_column = kpi_data
            starting_range = (self.starting_row, kpi_column)
            ending_range = (
                self.starting_row + number_of_brands,
                kpi_column + len(time_periods),
            )
            df = (
                sheet.range(starting_range, ending_range)
                .options(pd.DataFrame, index=False)
                .value
            )
            df.columns = ["brand"] + time_periods
            nuevas_columnas = ["brand"] + [self.reformatear_fecha(col) for col in df.columns[1:]] # ----> Código agregado
            df.columns = nuevas_columnas # ----> Código agregado
            #print(df.columns)
            #print(df.head())
            kpi = Kpi(name=kpi_name, data=df)
            self.kpis.append(kpi)

    def create_on_sheet(self, sheet: Sheet):
        sub_section_cell = sheet.range(f"A{self.target_row}")
        # Duplicate row
        #sub_section_is_brands = self.target_row == 5
        #if not sub_section_is_brands:
        self.duplicate_sub_section_row(sheet=sheet)
        # Write name of sub section
        sub_section_cell.value = self.name

    def duplicate_sub_section_row(self, sheet: Sheet):
        base_row = sheet.range("5:5")
        base_row.copy(sheet.range(f"{self.target_row}:{self.target_row}"))

    def get_target_row_number(self, index: int, brands_number: int):
        self.target_row = 5 + (3 + brands_number) * index

    def get_kpi_positions(self, sheet):
        for kpi in self.kpis:
            kpi.get_column_position(sheet=sheet)

    def copy_brands(self, sheet: Sheet):
        brands = self.kpis[0].data["brand"]
        brands_cell = sheet.range((self.target_row + 1, 1))
        
        brands_cell.options(index=False, header=False).value = brands

    def copy_data(self, sheet: Sheet, time_periods_number: int):
        for index, kpi in enumerate(self.kpis):
            kpi.copy_data(
                sheet=sheet,
                index=index,
                target_row=self.target_row,
                time_periods_number=time_periods_number,
            )
