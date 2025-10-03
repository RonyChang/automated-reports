import numpy as np
import pandas as pd
from xlwings import Sheet

from .kpi_equivalents import kpi_equivalents

class Kpi:
    def __init__(self, name: str, data: pd.DataFrame):
        self.powerview_name: str = name
        self.data: pd.DataFrame = data
        self.report_name: str = self.get_report_name()
        self.column_pos: int = None

    def get_report_name(self):
        return kpi_equivalents[self.powerview_name.strip()]

    def get_column_position(self, sheet: Sheet):
        cells_row = sheet.range("1:1").value
        cells = [
            (cell_name, index + 1)
            for index, cell_name in enumerate(cells_row)
            if cell_name is not None
        ]
        for cell in cells:
            if cell[0].strip() == self.report_name:
                self.column_pos = cell[1]
        return None

    def add_var_columns_to_df(
        self, first_difference: tuple[str, str], second_difference: tuple[str, str], third_difference: tuple[str, str]
    ):
        percentage_kpis = ["Share Volumen", "Share Valor", "Penetración"]
        if self.report_name in percentage_kpis:
            first_var = self.get_percentage_variation(difference=first_difference)
            second_var = self.get_percentage_variation(difference=second_difference)
            third_var = self.get_percentage_variation(difference=third_difference)
        else:
            first_var = self.get_normal_kpi_variation(difference=first_difference)
            second_var = self.get_normal_kpi_variation(difference=second_difference)
            third_var = self.get_normal_kpi_variation(difference=third_difference)
        self.data["first_var"] = first_var
        self.data["second_var"] = second_var
        self.data["third_var"] = third_var

    def get_percentage_variation(self, difference: tuple[str, str]):
        ''' Cambio 1'''
        df = self.data.copy()
        col1, col2 = difference
        # Convertir a numérico, "*" y cualquier texto raro se vuelven NaN
        s1 = pd.to_numeric(df[col1], errors="coerce")
        s2 = pd.to_numeric(df[col2], errors="coerce")
        # Evitar división por cero, reemplazamos 0 en la primera columna con NaN
        s1 = s1.replace(0, np.nan)
        # Calcular diferencia
        diff = s2 - s1
        # Donde haya NaN (porque había "*", texto o cero), devolver "-"
        diff = diff.where(diff.notna(), "-")
        return diff

    def get_normal_kpi_variation(self, difference: tuple[str, str]):
        ''' Cambio 2'''
        # Versión vectorizada: ((col2/col1) - 1) * 100, con manejo de 0 y NaN.
        df = self.data.copy()
        col1, col2 = difference
        s1 = pd.to_numeric(df[col1], errors='coerce')
        s2 = pd.to_numeric(df[col2], errors='coerce')
        s1 = s1.replace(0, np.nan)
        variation = ((s2 / s1) - 1) * 100
        return variation.where(pd.notna(variation), '-')

    def copy_data(
        self,
        sheet: Sheet,
        index: int,
        target_row: int,
        time_periods_number: int,
    ):
        if self.column_pos is None:
            print(f"\t\t\t-Column for {self.report_name} not found")
            return

        data_column = self.column_pos - time_periods_number - 3
        raw_data = self.data.drop(columns=["brand"])
        cell_position = sheet.range((target_row + 1, data_column))
        cell_position.options(index=False, header=False).value = raw_data
