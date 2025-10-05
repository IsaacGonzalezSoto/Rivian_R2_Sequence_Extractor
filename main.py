"""
Punto de entrada principal del sistema de extracción de L5X.
"""
import sys
import os
from src.pipeline.extraction_pipeline import ExtractionPipeline


def print_summary(routines_data):
    """
    Print a summary of extraction results.
    
    Args:
        routines_data: List of processed routines
    """
    print("\n" + "="*60)
    print("FINAL SUMMARY")
    print("="*60)
    
    for routine_info in routines_data:
        print(f"\n📄 {routine_info['routine_name']}")
        print(f"   Sequences:")
        print(f"     JSON: {routine_info.get('sequences_file', 'N/A')}")
        print(f"     CSV: {routine_info.get('sequences_csv', 'N/A')}")
        print(f"     Count: {routine_info.get('sequences_count', 0)}")
        
        if routine_info.get('transitions_file'):
            print(f"   Transitions:")
            print(f"     JSON: {routine_info['transitions_file']}")
            print(f"     CSV: {routine_info['transitions_csv']}")
            print(f"     Count: {routine_info['transitions_count']}")


def main():
    """
    Función principal del programa.
    """
    # Configuración
    l5x_file = "_010UA1_Fixture_Em0105_Program.L5X"
    output_folder = "output"
    debug = True
    
    # Validar que el archivo existe
    if not os.path.exists(l5x_file):
        print(f"❌ Error: No se encontró el archivo {l5x_file}")
        print("Por favor, actualiza la ruta del archivo L5X.")
        sys.exit(1)
    
    try:
        # Crear y ejecutar pipeline
        pipeline = ExtractionPipeline(
            l5x_file_path=l5x_file,
            output_folder=output_folder,
            debug=debug
        )
        
        # Ejecutar extracción
        routines_data = pipeline.run()
        
        # Mostrar resumen
        print_summary(routines_data)
        
        print("\n✅ Proceso completado exitosamente")
        
    except FileNotFoundError as e:
        print(f"❌ Error: Archivo no encontrado - {e}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Error durante el procesamiento: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()