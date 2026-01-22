"""
Script de emergencia para restaurar alertas en Excel.
Ejecutar si Excel quedó "congelado" con alertas desactivadas.
"""

import win32com.client as win32

def restaurar_alertas_excel():
    """Restaura alertas en TODAS las instancias de Excel abiertas."""
    
    print("="*60)
    print("🔧 RESTAURANDO ALERTAS EN EXCEL")
    print("="*60)
    
    instancias_encontradas = 0
    instancias_restauradas = 0
    
    try:
        # Intentar conectarse a instancias existentes de Excel
        import pythoncom
        
        # Obtener todas las instancias de Excel corriendo
        for i in range(10):  # Intentar hasta 10 instancias
            try:
                # GetActiveObject se conecta a instancia existente
                excel = win32.GetActiveObject("Excel.Application")
                
                if excel:
                    instancias_encontradas += 1
                    
                    print(f"\n📊 Instancia {instancias_encontradas} encontrada")
                    
                    # Restaurar configuración
                    try:
                        excel.DisplayAlerts = True
                        excel.ScreenUpdating = True
                        excel.Interactive = True
                        excel.EnableEvents = True
                        
                        # Mostrar info
                        num_workbooks = excel.Workbooks.Count
                        print(f"   ✅ Alertas restauradas")
                        print(f"   📁 Archivos abiertos: {num_workbooks}")
                        
                        if num_workbooks > 0:
                            print("   📋 Archivos:")
                            for j in range(1, min(num_workbooks + 1, 6)):
                                try:
                                    wb_name = excel.Workbooks(j).Name
                                    print(f"      - {wb_name}")
                                except:
                                    pass
                        
                        instancias_restauradas += 1
                        
                    except Exception as e:
                        print(f"   ⚠️ Error al restaurar: {e}")
                
                break  # Si encontramos una, salir del loop
                
            except pythoncom.com_error:
                # No hay más instancias
                break
            except Exception as e:
                print(f"   ⚠️ Error: {e}")
                break
        
        # Si no encontramos ninguna instancia
        if instancias_encontradas == 0:
            print("\nℹ️  No se encontraron instancias de Excel abiertas")
            print("   Si Excel está abierto, intenta con la Opción 2 (VBA)")
        else:
            print("\n" + "="*60)
            print(f"✅ PROCESO COMPLETADO")
            print(f"   Instancias encontradas: {instancias_encontradas}")
            print(f"   Instancias restauradas: {instancias_restauradas}")
            print("="*60)
    
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        print("\n💡 SOLUCIÓN ALTERNATIVA:")
        print("   1. Cierra todos los archivos Excel")
        print("   2. Cierra Excel completamente")
        print("   3. Vuelve a abrir Excel")
    
    print("\n✅ Presiona ENTER para salir...")
    input()


if __name__ == "__main__":
    restaurar_alertas_excel()