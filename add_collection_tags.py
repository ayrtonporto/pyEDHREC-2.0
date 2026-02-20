import pandas as pd
import re


# =========================
# CONFIG
# =========================
INPUT_FILE = "lista_cartas.txt"
OUTPUT_FILE = "lista_cartas_con_tags.txt"
INVENTORY_FILE = "inventario.csv"


# =========================
# 1. CARGAR INVENTARIO
# =========================
def load_inventory():
    """
    Carga el inventario y crea un mapa de carta -> colecciones
    """
    print(f"📦 Cargando inventario desde: {INVENTORY_FILE}")
    
    inv = pd.read_csv(INVENTORY_FILE)
    inv["name_lower"] = inv["name"].str.lower()
    
    # Crear mapa: nombre_carta -> lista de colecciones
    card_collections = {}
    
    for _, row in inv.iterrows():
        name_lower = row["name_lower"]
        collection = str(row["source"])
        
        if name_lower not in card_collections:
            card_collections[name_lower] = []
        
        if collection not in card_collections[name_lower]:
            card_collections[name_lower].append(collection)
    
    print(f"   ✅ {len(card_collections)} cartas únicas cargadas")
    return card_collections


# =========================
# 2. SANITIZAR NOMBRE DE COLECCIÓN
# =========================
def sanitize_collection_name(collection):
    """
    Convierte nombre de colección a formato tag limpio
    Ejemplo: "Commander Precons" -> "Commander_Precons"
    """
    # Remover caracteres especiales
    clean = collection.replace(",", "").replace("'", "").replace(":", "")
    clean = clean.replace("/", "_").replace("-", "_")
    # Reemplazar espacios con guiones bajos
    clean = "_".join(clean.split())
    return clean


# =========================
# 3. LIMPIAR NOMBRE DE CARTA
# =========================
def clean_card_name(line):
    """
    Limpia una línea para extraer el nombre de la carta
    Remueve:
    - Números al inicio (1 Sol Ring → Sol Ring)
    - Tags existentes (#Commander)
    - Marcadores (⭐ 🔥)
    - Comentarios (#)
    """
    # Remover comentarios completos
    if line.strip().startswith("#"):
        return None, line.strip()
    
    # Guardar la línea original
    original = line.strip()
    
    # Remover número al inicio
    cleaned = re.sub(r'^\d+\s+', '', line)
    
    # Remover tags existentes
    cleaned = re.sub(r'#\S+', '', cleaned)
    
    # Remover marcadores emoji
    cleaned = re.sub(r'[⭐🔥💰🔗📘📗\[\]]', '', cleaned)
    
    # Remover espacios extra
    cleaned = cleaned.strip()
    
    return cleaned, original


# =========================
# 4. PROCESAR ARCHIVO
# =========================
def process_file(input_file, output_file, card_collections):
    """
    Lee el archivo de entrada y agrega tags de colección
    """
    print(f"\n📄 Procesando archivo: {input_file}")
    
    found_count = 0
    not_found_count = 0
    comment_count = 0
    
    output_lines = []
    not_found_cards = []
    
    with open(input_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    for line in lines:
        # Limpiar la carta
        card_name, original_line = clean_card_name(line)
        
        # Si es comentario completo, mantener tal cual
        if card_name is None:
            output_lines.append(original_line)
            comment_count += 1
            continue
        
        # Si es línea vacía, mantener
        if not card_name:
            output_lines.append("")
            continue
        
        # Buscar en inventario
        card_lower = card_name.lower()
        
        if card_lower in card_collections:
            collections = card_collections[card_lower]
            
            # Si tiene múltiples colecciones, usar la primera
            collection = collections[0]
            
            # Sanitizar nombre de colección
            collection_tag = sanitize_collection_name(collection)
            
            # Reconstruir línea
            # Detectar si tenía número al inicio
            match = re.match(r'^(\d+\s+)', line)
            if match:
                prefix = match.group(1)
                new_line = f"{prefix}{card_name} #{collection_tag}"
            else:
                new_line = f"{card_name} #{collection_tag}"
            
            output_lines.append(new_line)
            found_count += 1
        else:
            # Carta no encontrada en inventario
            output_lines.append(original_line + " #NOT_FOUND")
            not_found_cards.append(card_name)
            not_found_count += 1
    
    # Escribir archivo de salida
    with open(output_file, 'w', encoding='utf-8') as f:
        for line in output_lines:
            f.write(line + "\n")
    
    # Mostrar estadísticas
    print(f"\n✅ Procesamiento completado:")
    print(f"   • {found_count} cartas encontradas y taggeadas")
    print(f"   • {not_found_count} cartas NO encontradas en inventario")
    print(f"   • {comment_count} líneas de comentario preservadas")
    print(f"\n📁 Archivo generado: {output_file}")
    
    # Mostrar cartas no encontradas
    if not_found_cards:
        print(f"\n⚠️  Cartas NO encontradas en tu inventario:")
        for card in not_found_cards[:10]:
            print(f"   - {card}")
        if len(not_found_cards) > 10:
            print(f"   ... y {len(not_found_cards) - 10} más")


# =========================
# 5. PROCESAR MÚLTIPLES COLECCIONES (OPCIONAL)
# =========================
def process_with_all_collections(input_file, output_file, card_collections):
    """
    Versión alternativa que muestra TODAS las colecciones donde aparece la carta
    """
    print(f"\n📄 Procesando archivo (modo todas las colecciones): {input_file}")
    
    found_count = 0
    not_found_count = 0
    
    output_lines = []
    
    with open(input_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    for line in lines:
        card_name, original_line = clean_card_name(line)
        
        if card_name is None:
            output_lines.append(original_line)
            continue
        
        if not card_name:
            output_lines.append("")
            continue
        
        card_lower = card_name.lower()
        
        if card_lower in card_collections:
            collections = card_collections[card_lower]
            
            # Crear tags para todas las colecciones
            tags = " ".join([f"#{sanitize_collection_name(c)}" for c in collections])
            
            # Reconstruir línea
            match = re.match(r'^(\d+\s+)', line)
            if match:
                prefix = match.group(1)
                new_line = f"{prefix}{card_name} {tags}"
            else:
                new_line = f"{card_name} {tags}"
            
            output_lines.append(new_line)
            found_count += 1
        else:
            output_lines.append(original_line + " #NOT_FOUND")
            not_found_count += 1
    
    # Escribir archivo
    with open(output_file, 'w', encoding='utf-8') as f:
        for line in output_lines:
            f.write(line + "\n")
    
    print(f"\n✅ Procesamiento completado:")
    print(f"   • {found_count} cartas taggeadas")
    print(f"   • {not_found_count} cartas NO encontradas")
    print(f"\n📁 Archivo generado: {output_file}")


# =========================
# MAIN
# =========================
if __name__ == "__main__":
    print("="*70)
    print("🏷️  AGREGADOR DE TAGS DE COLECCIÓN")
    print("="*70)
    
    # Cargar inventario
    card_collections = load_inventory()
    
    # Verificar que existe el archivo de entrada
    import os
    if not os.path.exists(INPUT_FILE):
        print(f"\n❌ Error: No se encuentra el archivo '{INPUT_FILE}'")
        print(f"   Crea un archivo con ese nombre con tu lista de cartas.")
        print(f"\n   Formato esperado:")
        print(f"   1 Sol Ring")
        print(f"   1 Arcane Signet")
        print(f"   1 Rhystic Study")
        exit(1)
    
    # Preguntar modo
    print(f"\n🔧 Opciones de procesamiento:")
    print(f"   1. Una colección por carta (primera encontrada)")
    print(f"   2. Todas las colecciones por carta")
    
    choice = input("\n👉 Elige opción (1 o 2, default=1): ").strip()
    
    if choice == "2":
        # Modo: todas las colecciones
        output_file = OUTPUT_FILE.replace(".txt", "_todas_colecciones.txt")
        process_with_all_collections(INPUT_FILE, output_file, card_collections)
    else:
        # Modo: una colección (default)
        process_file(INPUT_FILE, OUTPUT_FILE, card_collections)
    
    print("\n" + "="*70)
    print("✅ ¡Listo!")
    print("="*70)
