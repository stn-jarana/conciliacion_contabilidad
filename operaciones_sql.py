from conexion import get_connection

def obtener_tipo_cambio(fecha):
    conexion = get_connection()

    try:

        cursor = conexion.cursor()
        sql = "SELECT Tipo_Venta FROM dbo.CN_TipoCambio WHERE fecha = ?"
        cursor.execute(sql, fecha)

        return cursor.fetchall()

    except Exception as e:
        print(e)

    finally:
        conexion.close()


def import_xlsm_sql():
    conexion = get_connection()

    try:
        cursor = conexion.cursor()
        sql = ""
        cursor.execute(sql, )
        pass

    except Exception as e:
        pass
    finally:
        conexion.close()
    pass
