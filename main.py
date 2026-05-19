import pandas as pd

# Leer el archivo CSV
df = pd.read_csv("datos.csv", encoding="utf-8-sig")

# Limpiar nombres de columnas
df.columns = df.columns.str.strip().str.lower()

while True:
    print("\n=== MENÚ ===")
    print("1. Mostrar nombres")
    print("2. Mostrar edades")
    print("3. Mostrar sexo")
    print("4. Nombre y edad")
    print("5. Nombre y sexo")
    print("6. Edad y sexo")
    print("7. Mostrar todo")
    print("0. Salir")

    opcion = input("\nElige una opción (0-7): ")

    if opcion == "0":
        print("👋 Saliendo del programa...")
        break

    print("\n=== RESULTADO ===")

    try:
        if opcion == "1":
            print(df["nombre"].to_string(index=False))

        elif opcion == "2":
            print(df["edad"].to_string(index=False))

        elif opcion == "3":
            print(df["sexo"].to_string(index=False))

        elif opcion == "4":
            print(df[["nombre", "edad"]].to_string(index=False))

        elif opcion == "5":
            print(df[["nombre", "sexo"]].to_string(index=False))

        elif opcion == "6":
            print(df[["edad", "sexo"]].to_string(index=False))

        elif opcion == "7":
            print(df.to_string(index=False))

        else:
            print("❌ Opción no válida")

    except KeyError as e:
        print("❌ Error: columna no encontrada ->", e)
        print("Columnas disponibles:", df.columns.tolist())