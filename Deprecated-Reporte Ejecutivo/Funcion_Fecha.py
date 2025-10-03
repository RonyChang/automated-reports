from datetime import datetime, timedelta

def obtener_fecha_formateada():
    # Obtener la fecha actual
    fecha_actual = datetime.now()

    # Restar un mes
    mes_anterior = fecha_actual.replace(day=1) - timedelta(days=1)

    # Formatear la fecha como "Mes-Año"
    fecha_formateada = mes_anterior.strftime("%b-%y").capitalize()

    return fecha_formateada 