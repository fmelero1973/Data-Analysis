import tableauserverclient as TSC
import os

def exportar_vista_filtrada(server_url, token_name, token_secret, site_id,
                             workbook_name, vista_nombre, filtros, ruta_destino):
    """
    Exporta una imagen PNG de una vista de Tableau con filtros aplicados.

    filtros: diccionario {nombre_filtro: valor_filtro}
    ruta_destino: ruta donde guardar la imagen, incluyendo nombre de archivo
    """
    auth = TSC.PersonalAccessTokenAuth(token_name, token_secret, site_id)
    server = TSC.Server(server_url, use_server_version=True)

    with server.auth.sign_in(auth):
        all_views, _ = server.views.get()
        all_workbooks, _ = server.workbooks.get()
        workbook_id_to_name = {wb.id: wb.name for wb in all_workbooks}

        # Buscar la vista
        vista = next(
            (v for v in all_views if v.name == vista_nombre and
             workbook_id_to_name.get(v.workbook_id) == workbook_name),
            None
        )

        if not vista:
            print(f"Vista '{vista_nombre}' en '{workbook_name}' no encontrada.")
            return None

        # Aplicar filtros como dict: {"filtro": "valor"}
        req_options = TSC.RequestOptions()
        req_options.vf = filtros  # visualización filtrada

        # Descargar imagen con filtros aplicados
        image_bytes = server.views.populate_image(vista, req_options).image

        # Guardar imagen
        with open(ruta_destino, "wb") as f:
            f.write(image_bytes)

        print(f"Vista '{vista_nombre}' exportada como imagen en: {ruta_destino}")
        return ruta_destino