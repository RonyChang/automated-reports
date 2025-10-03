import time
import xlwings as xw
import yaml

from classes.preprocessing import PreProcessing
from classes.report import Report

with open("config/config.yaml", "r", encoding="utf-8") as f:
    _cfg = yaml.safe_load(f)

pp = PreProcessing(
    _cfg["file_in"],
    _cfg["file_template"],
    _cfg["file_out"],
)
pp.run()

NAME = _cfg["name"]
PERIOD = _cfg["period"]
FIRST_DIFFERENCE, SECOND_DIFFERENCE, THIRD_DIFFERENCE = [
    tuple(x) for x in _cfg.get("differences", [(), (), ()])
]

def main():
    report = Report(
        name=NAME,
        period=PERIOD,
        first_difference=FIRST_DIFFERENCE,
        second_difference=SECOND_DIFFERENCE,
        third_difference=THIRD_DIFFERENCE,
    )
    start_time = time.time()
    report.create_sections()
    report.start_sections()

    report.get_data()
    report.insert_columns_for_time_periods()
    report.create_kpis_sections()
    report.create_sub_sections_on_sheet()
    report.get_kpi_positions()
    report.add_var_columns_to_df()

    report.copy_data()
    report.copy_brands()
    report.style_report()
    report.add_hyperlinks()

    report.final_styling()
    report.clean_subsections()
    report.remove_first_rows()
    report.apply_star_bold_underline(_cfg["file_template"])
    report.format_two_decimals_and_zero_dash()
    report.paint_only_kpi_columns_post()

    report.save()
    end_time = time.time()

    # Calculate elapsed time in seconds
    elapsed_time = end_time - start_time

    # Convert seconds to minutes
    elapsed_minutes = round(elapsed_time / 60, 2)

    print(f"Tiempo transcurrido: {elapsed_minutes} minutos")

if __name__ == "__main__":
    with xw.App(visible=False) as _:
        main()
