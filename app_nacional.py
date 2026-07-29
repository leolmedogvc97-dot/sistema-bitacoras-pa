with col_acc2:
            if st.button("🚀 Guardar y Generar 3 Bitácoras Definitivas"):
                try:
                    wb = openpyxl.load_workbook("Prueba unificación.xlsx")
                    ws = wb["BASE_DE_DATOS"]
                    
                    # Limpiar filas previas para evitar encimar datos
                    if ws.max_row >= 2:
                        ws.delete_rows(2, ws.max_row)
                    
                    # 1. Escribir nuevos encabezados para las columnas adicionales
                    ws["K1"] = "Rendimiento (km/L)"
                    ws["L1"] = "Precio Gasolina ($)"
                    ws["M1"] = "GASTO COMBUSTI"
                    
                    for i, reg in enumerate(registros_totales):
                        row_idx = 2 + i
                        
                        # 2. Agregar los datos iterados con las fórmulas de Excel
                        fila_datos = [
                            reg["FECHA COMPLETA"],
                            f'=UPPER(TEXT(A{row_idx}, "MMMM"))',
                            reg["MUNICIPIO"],
                            reg["POBLADO"],
                            reg["folio CIIA"],
                            reg["HORA DE SALIDA"],
                            reg["KM INICIAL / Km de Salida"],
                            reg["RECORRIDO"],                       # Columna H (Kilometraje)
                            reg["HORA DE LLEGADA"],
                            reg["KM FINAL / Km de Llegada"],
                            12.0,                                   # Columna K: Rendimiento (Dato base)
                            23.99,                                  # Columna L: Precio de gasolina (Dato base)
                            f'=ROUND((H{row_idx}/K{row_idx})*L{row_idx}, 2)', # Columna M: FÓRMULA DE GASTO
                            reg["Gasolina de Salida"],
                            reg["Gasolina de Llegada"],
                            reg["Dotación de Gasolina(LLENAR GASTO DE COMBUSTIBLE)"],
                            reg["Oficio Numero"],
                            reg["Oficio Fecha"],
                            reg["observaciones"],
                            reg["Usuario Responsable"],
                            reg["Áreas de Adscripción"],
                            reg["Tipo de Vehículo"],
                            reg["Placas"],
                            reg["No. De Licencia"]
                        ]
                        ws.append(fila_datos)
                    
                    # 3. CREAR LA FILA DE TOTALES
                    last_row = ws.max_row + 1
                    ws[f'A{last_row}'] = "TOTALES"
                    ws[f'H{last_row}'] = f'=SUM(H2:H{last_row-1})' # Suma de Recorrido
                    ws[f'M{last_row}'] = f'=SUM(M2:M{last_row-1})' # Suma de Gasto de Combustible
                    
                    # 4. DAR FORMATO ROJO A LA FILA DE TOTALES
                    from openpyxl.styles import PatternFill, Font
                    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                    white_font = Font(color="FFFFFF", bold=True)
                    
                    # Colorear toda la fila de totales de rojo con letra blanca
                    for col in range(1, 25):  # De la columna A hasta la X
                        cell = ws.cell(row=last_row, column=col)
                        cell.fill = red_fill
                        if col in [1, 8, 13]: # Remarcar texto de totales en negritas (A, H, M)
                            cell.font = white_font
                    
                    output = BytesIO()
                    wb.save(output)
                    output.seek(0)
                    
                    registrar_auditoria("GENERAR BITACORAS", "Generación y descarga de archivo unificado con fórmulas y totales")
                    st.success("¡Archivo generado con éxito y listo para descarga!")
                    st.download_button(
                        label="⬇️ Descargar Archivo Definitivo (Incluye las 3 Bitácoras)",
                        data=output,
                        file_name="BITACORAS_OFICIALES_DEFINITIVAS.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"Error al generar el archivo definitivo: {e}")
