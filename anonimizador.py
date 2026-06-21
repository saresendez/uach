import pandas as pd
import re
import os

def extract_colonia(direccion):
    if pd.isna(direccion): return direccion
    # Extrae solo la parte de la dirección que indica Colonia o Fraccionamiento
    match = re.search(r'(COL\s.*|FRACC\s.*)', str(direccion), re.IGNORECASE)
    return match.group(1).strip() if match else direccion

def procesar_anonimizacion(input_file):
    if not os.path.exists(input_file):
        print(f"Error: No se encuentra el archivo {input_file} en {os.getcwd()}")
        return

    # Cargar datos
    df = pd.read_excel(input_file)
    
    # 1. Crear ID único (CLI-0001 en adelante)
    df['ID_CLIENTE'] = [f'CLI-{i:04d}' for i in range(1, len(df) + 1)]
    
    # 2. Generar el archivo de Mapeo
    mapeo = df[['ID_CLIENTE', 'RFC', 'NOMBRE', 'CORREO', 'TELEFONO']].copy()
    # Modificado: Se cambia .to_csv por .to_excel
    mapeo.to_excel('mapeo_identidad.xlsx', index=False)
    
    # 3. Anonimización para el archivo de trabajo
    df['RFC'] = 'xxxxxxxxxxxxx'
    df['TELEFONO'] = 'xxx-xxx-xxxx'
    df['CORREO'] = [f'anon{i:04d}@anon.com' for i in range(1, len(df) + 1)]
    df['DIRECCION'] = df['DIRECCION'].apply(extract_colonia)
    
    # Eliminar todas las columnas que contengan 'NOMBRE'
    cols_to_drop = [c for c in df.columns if 'NOMBRE' in str(c).upper()]
    df = df.drop(columns=cols_to_drop)
    
    # 4. Ordenar: ID_CLIENTE primero, luego lo demás
    cols = ['ID_CLIENTE'] + [c for c in df.columns if c != 'ID_CLIENTE']
    df = df[cols]
    
    # Exportar
    # Modificado: Se cambia .to_csv por .to_excel
    df.to_excel('reg_anon.xlsx', index=False)
    
    print(f"✅ Proceso completado exitosamente.")
    print(f"📦 Se han procesado {len(df)} registros.")
    print(f"📄 Archivos generados:")
    print(f"   - reg_anon.xlsx (Para compartir)")
    print(f"   - mapeo_identidad.xlsx (TU ARCHIVO PRIVADO - NO COMPARTIR)")

if __name__ == "__main__":
    procesar_anonimizacion('registros.xlsx')