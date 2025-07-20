import pandas as pd
from datetime import datetime, timedelta, time
from google.colab import files

# --- FUNCIONES BASE ---
def convertir_a_str(hora):
    if isinstance(hora, time):
        return hora.strftime("%H:%M:%S")
    elif isinstance(hora, str):
        return hora
    return None

def calcular_horas(inicio_raw, fin_raw, refrigerio_inicio_raw=None, refrigerio_fin_raw=None):
    formato = "%H:%M:%S"
    inicio_str = convertir_a_str(inicio_raw)
    fin_str = convertir_a_str(fin_raw)
    refrigerio_inicio_str = convertir_a_str(refrigerio_inicio_raw)
    refrigerio_fin_str = convertir_a_str(refrigerio_fin_raw)

    if not inicio_str or not fin_str:
        return [0]*8

    try:
        inicio = datetime.strptime(inicio_str, formato)
        fin = datetime.strptime(fin_str, formato)
        if fin <= inicio:
            fin += timedelta(days=1)

        minutos_refrigerio = 0
        if refrigerio_inicio_str and refrigerio_fin_str:
            ri = datetime.strptime(refrigerio_inicio_str, formato).time()
            rf = datetime.strptime(refrigerio_fin_str, formato).time()
            if ri == time(13, 0) and rf == time(14, 0):
                minutos_refrigerio = 60
            elif ri == time(12, 0) and rf == time(12, 45):
                minutos_refrigerio = 45

        minutos_diurnos_total = 0
        minutos_nocturnos_total = 0

        actual = inicio
        while actual < fin:
            hora = actual.time()
            if time(6, 0) <= hora < time(22, 0):
                minutos_diurnos_total += 1
            else:
                minutos_nocturnos_total += 1
            actual += timedelta(minutes=1)

        total_minutos = minutos_diurnos_total + minutos_nocturnos_total

        if minutos_refrigerio > 0:
            if minutos_diurnos_total >= minutos_refrigerio:
                minutos_diurnos_total -= minutos_refrigerio
            else:
                restante = minutos_refrigerio - minutos_diurnos_total
                minutos_diurnos_total = 0
                minutos_nocturnos_total = max(0, minutos_nocturnos_total - restante)
            total_minutos -= minutos_refrigerio

        minutos_normales = min(total_minutos, 480)
        minutos_extras = max(0, total_minutos - 480)

        minutos_diurnos_normales = 0
        minutos_nocturnos_normales = 0
        actual = inicio
        minutos_asignados = 0
        while actual < fin and minutos_asignados < minutos_normales:
            hora = actual.time()
            if time(6, 0) <= hora < time(22, 0):
                minutos_diurnos_normales += 1
            else:
                minutos_nocturnos_normales += 1
            minutos_asignados += 1
            actual += timedelta(minutes=1)

        horas_diurnas = minutos_diurnos_normales / 60
        horas_nocturnas = minutos_nocturnos_normales / 60

        horas_extra_25 = min(minutos_extras, 120) / 60
        horas_extra_35 = max(minutos_extras - 120, 0) / 60

        horas_extra_25_nocturna = 0
        horas_extra_35_nocturna = 0

        if inicio.time() >= time(15, 0) and inicio.time() < time(20, 0):
            horas_extra_25_nocturna = horas_extra_25
            horas_extra_35_nocturna = round(horas_diurnas, 2) - horas_extra_25_nocturna
            horas_extra_25 = 0
            horas_extra_35 = horas_extra_35 - horas_extra_35_nocturna

        if inicio.time() >= time(20, 0) and inicio.time() < time(22, 0):
            horas_extra_25_nocturna = round(horas_diurnas, 2)
            horas_extra_35_nocturna = 0
            horas_extra_25 = horas_extra_25 - horas_extra_25_nocturna

        if fin.time() >= time(22, 0):
            horas_extra_35_nocturna = ((fin - datetime.combine(fin.date(), time(22, 0))).seconds / 60)/60
            horas_extra_35 = horas_extra_35-horas_extra_35_nocturna

        if fin.time() < time(6, 0):
            inicio = datetime.combine(fin.date(), time(22, 0)) - timedelta(days=1)  # 10:00 PM del día anterior
            diferencia = fin - inicio
            horas_extra_35_nocturna = (diferencia.seconds / 60) / 60
            horas_extra_35 = horas_extra_35-horas_extra_35_nocturna

        total_horas = (minutos_diurnos_total + minutos_nocturnos_total) / 60

        return max(round(horas_diurnas, 2), 0), max(round(horas_nocturnas, 2), 0), \
               max(round(minutos_normales / 60, 2), 0), max(round(horas_extra_25, 2), 0), \
               max(round(horas_extra_35, 2), 0), max(round(horas_extra_25_nocturna, 2), 0), \
               max(round(horas_extra_35_nocturna, 2), 0), max(round(total_horas, 2), 0)

    except Exception:
        return [0]*8

# --- PROCESAR FILA ---
def procesar_fila(row):
    resultado = calcular_horas(
        row["Hora Inicio Labores"],
        row["Hora Término Labores"],
        row.get("Hora Inicio Refrigerio", None),
        row.get("Hora Término Refrigerio", None)
    )

    dia = str(row["DIA"]).strip().lower()
    dias_normales = ['lunes', 'martes', 'miércoles', 'miercoles', 'jueves', 'viernes', 'sábado', 'sabado']

    if dia in dias_normales:
        return pd.Series({
            "Horas Diurnas": resultado[0],
            "Extra 25%": resultado[3],
            "Extra 35%": resultado[4],
            "Horas Nocturnas": resultado[1],
            "Extra 25% Nocturna": resultado[5],
            "Extra 35% Nocturna": resultado[6],
            "Horas Domingo/Feriado": 0,
            "Horas Extra Domingo/Feriado": 0,
            "Horas Normales": resultado[2],
            "Total Horas": resultado[7],
        })
    else:
        total_horas = resultado[7]
        base = min(total_horas, 8)
        extra = max(total_horas - 8, 0)
        return pd.Series({
            "Horas Diurnas": 0,
            "Extra 25%": 0,
            "Extra 35%": 0,
            "Horas Nocturnas": 0,
            "Extra 25% Nocturna": 0,
            "Extra 35% Nocturna": 0,
            "Horas Domingo/Feriado": round(base, 2),
            "Horas Extra Domingo/Feriado": round(extra, 2),
            "Horas Normales": 0,
            "Total Horas": total_horas,
        })

# --- CALCULAR DIA-TRA Y MODIFICAR SI ES DESCANSO MÉDICO ---
def calcular_dia_tra(row):
    dia = row["DIA"].strip().lower()
    es_laboral = dia in ["lunes", "martes", "miércoles", "miercoles", "jueves", "viernes", "sábado", "sabado"]
    tiene_horas = row["Horas Diurnas"] + row["Horas Nocturnas"] > 0
    
    # Condición para Descanso Médico
    if row["Labor/Actividad"] == "Descanso Médico":
        return "DM"
    
    return 1 if es_laboral and tiene_horas else 0

# --- EJECUCIÓN PRINCIPAL ---
uploaded = files.upload()  # Cargar archivo desde tu máquina local
file_name = next(iter(uploaded))  # Obtener el nombre del archivo cargado
df = pd.read_excel(file_name)  # Leer el archivo Excel

columnas_ingreso = df.columns.tolist()

# Aplicar el procesamiento de filas
resultados = df.apply(procesar_fila, axis=1, result_type="expand")
df_resultado = pd.concat([df, resultados], axis=1)

# Modificar la columna "DIA-TRA"
df_resultado["DIA-TRA"] = df_resultado.apply(calcular_dia_tra, axis=1)

# --- ORDEN FINAL DE COLUMNAS (según tu imagen) ---
columnas_orden = [
    "Horas Diurnas", "Extra 25%", "Extra 35%", "Horas Nocturnas",
    "Extra 25% Nocturna", "Extra 35% Nocturna",
    "Horas Domingo/Feriado", "Horas Extra Domingo/Feriado",
    "Horas Normales", "Total Horas"
]
df_resultado = df_resultado[columnas_ingreso + ["DIA-TRA"] + columnas_orden]

# --- GUARDAR ARCHIVO FINAL ---
df_resultado.to_excel("reporte_horas_final_actualizado(C).xlsx", index=False)
print("✅ Archivo generado: 'reporte_horas_final_actualizado(C).xlsx'")

# Descargar el archivo con los resultados
files.download("reporte_horas_final_actualizado(C).xlsx")
