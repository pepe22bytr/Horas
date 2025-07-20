from horas_laborales.procesador import procesar_archivo

if __name__ == "__main__":
    archivo = "data/ejemplo_input.xlsx"
    procesar_archivo(archivo, "reporte_horas_final_actualizado.xlsx")
    print("✅ Reporte generado correctamente.")
