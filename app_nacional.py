with col_acc2:
            if st.button("🚀 Guardar y Generar 3 Bitácoras Definitivas"):
                try:
                    wb = openpyxl.load_workbook("Prueba unificación.xlsx")
                    ws = wb["BASE_DE_DATOS"]
                    
                    # 1. Limpiar filas previas manteniendo intactas las demás hojas
                    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=25):
                        for cell in row:
                            cell.value = None
                            
                    # 2. Encabezados adicionales
                    ws["K1"] = "Rendimiento (km/L)"
                    ws["L1"] = "Precio Gasolina ($)"
                    ws["M1"] = "GASTO COMBUSTI"
                    
                    # 3. Escribir los datos en el orden exacto de columnas que esperan las otras hojas
                    for i, reg in enumerate(registros_totales):
                        row_idx = 2 + i
                        
                        ws[f'A{row_idx}'] = reg["FECHA COMPLETA"]                                   # A: Fecha
                        ws[f'B{row_idx}'] = f'=UPPER(TEXT(A{row_idx}, "MMMM"))'                   # B: Mes
                        ws[f'C{row_idx}'] = reg["MUNICIPIO"]                                      # C: Municipio
                        ws[f'D{row_idx}'] = reg["POBLADO"]                                        # D: Poblado
                        ws[f'E{row_idx}'] = reg["folio CIIA"]                                     # E: Folio CIIA
                        ws[f'F{row_idx}'] = reg["HORA DE SALIDA"]                                 # F: Hora Salida
                        ws[f'G{row_idx}'] = reg["KM INICIAL / Km de Salida"]                      # G: Km Salida
                        ws[f'H{row_idx}'] = reg["RECORRIDO"]                                      # H: Recorrido (Km)
                        ws[f'I{row_idx}'] = reg["HORA DE LLEGADA"]                                # I: Hora Llegada
                        ws[f'J{row_idx}'] = reg["KM FINAL / Km de Llegada"]                       # J: Km Llegada
                        ws[f'K{row_idx}'] = 12.0                                                  # K: Rendimiento
                        ws[f'L{row_idx}'] = 23.99                                                 # L: Precio Gasolina
                        ws[f'M{row_idx}'] = f'=ROUND((H{row_idx}/K{row_idx})*L{row_idx}, 2)'        # M: Fórmula Gasto
                        ws[f'N{row_idx}'] = reg["Gasolina de Salida"]                             # N: Gas Salida
                        ws[f'O{row_idx}'] = reg["Oficio Numero"]                                  # O: Oficio Número
                        ws[f'P{row_idx}'] = reg["Oficio Fecha"]                                   # P: Oficio Fecha
                        ws[f'Q{row_idx}'] = reg["observaciones"]                                  # Q: Observaciones
                        ws[f'R{row_idx}'] = reg["Usuario Responsable"]                            # R: Usuario
                        ws[f'S{row_idx}'] = reg["Áreas de Adscripción"]                           # S: Adscripción
                        ws[f'T{row_idx}'] = reg["Tipo de Vehículo"]                               # T: Tipo Vehículo
                        ws[f'U{row_idx}'] = reg["Placas"]                                         # U: Placas
                        ws[f'V{row_idx}'] = reg["No. De Licencia"]                                # V: Licencia

                    # 4. Guardar archivo y descargar
                    output = BytesIO()
                    wb.save(output)
                    output.seek(0)
                    
                    registrar_auditoria("GENERAR BITACORAS", "Generación exitosa con mapeo correcto de columnas")
                    st.success("¡Archivo generado con éxito y listo para descarga!")
                    st.download_button(
                        label="⬇️ Descargar Archivo Definitivo (Incluye las 3 Bitácoras)",
                        data=output,
                        file_name="BITACORAS_OFICIALES_DEFINITIVAS.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"Error al generar el archivo definitivo: {e}")
