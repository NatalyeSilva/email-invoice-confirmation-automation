import pandas as pd
import os
import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog
from datetime import datetime
import win32com.client as win32

# =====================================================
# === CONFIGURACIÓN GENERAL ===
# =====================================================

archivo_excel = 'Recepción.xlsx'
hoja = 'REGISTROS'
columna_fecha = 'Fecha Pago Cargado Tesoreria'
columnas_deseadas = [
    'Empresa', 'Proveedor', 'Factura', 'Monto', 'Moneda', 'Fecha Pago Cargado Tesoreria'
]
empresa_filtrar = 'Nutrex_INC'
archivo_salida = 'datos_filtrados_nutrex.xlsx'

archivo_remitentes = 'remitentes.xlsx'
columna_proveedor = 'Proveedor'
columna_para = 'PARA'
columna_cc = 'CC'

ruta_adjuntos = r'C:\Users\Nutrex\OneDrive - NUTREX PARAGUAY SRL\Documentos - Documentos\ADMINISTRACION\25.BANCOS\Bancos BVI\Bancos Actuales\1-Bancolombia\Transferencias\2026'


# =====================================================
# === FUNCIONES ===
# =====================================================

def log(mensaje):
    """Agrega un mensaje al área de comentarios de la ventana."""
    txt_log.configure(state='normal')
    txt_log.insert(tk.END, mensaje + '\n')
    txt_log.see(tk.END)
    txt_log.configure(state='disabled')
    ventana.update_idletasks()


def limpiar_log():
    """Limpia el área de comentarios."""
    txt_log.configure(state='normal')
    txt_log.delete('1.0', tk.END)
    txt_log.configure(state='disabled')


def generar_archivo():
    """PASO 1: Filtra datos por fecha y genera el archivo Excel."""
    fechas_texto = entrada_fecha.get().strip()

    # Validar fechas
    if fechas_texto != '':
        for f in fechas_texto.split(','):
            try:
                datetime.strptime(f.strip(), '%Y-%m-%d')
            except ValueError:
                messagebox.showwarning("Formato inválido", f"Fecha no válida: '{f.strip()}'\nUsa formato YYYY-MM-DD")
                return

    limpiar_log()
    log("=" * 50)
    log("  PASO 1: GENERANDO ARCHIVO FILTRADO")
    log("=" * 50)

    try:
        df = pd.read_excel(archivo_excel, sheet_name=hoja)
        df.columns = df.columns.str.strip()
        df = df[df['Empresa'] == empresa_filtrar]
        df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce')

        if fechas_texto.strip() == '':
            fechas_filtrar = [df[columna_fecha].max().date()]
        else:
            fechas_filtrar = [
                datetime.strptime(f.strip(), '%Y-%m-%d').date()
                for f in fechas_texto.split(',')
            ]

        df_filtrado = df[df[columna_fecha].dt.date.isin(fechas_filtrar)]

        if df_filtrado.empty:
            log(f"⚠️ No hay datos para las fechas: {', '.join(str(f) for f in fechas_filtrar)}")
            messagebox.showwarning("Sin datos", f"No hay datos para las fechas: {', '.join(str(f) for f in fechas_filtrar)}")
            return

        df_filtrado = df_filtrado[columnas_deseadas]
        df_filtrado.to_excel(archivo_salida, index=False)

        log(f"\n📅 Fechas filtradas: {', '.join(str(f) for f in fechas_filtrar)}")
        log(f"📄 Archivo generado: {archivo_salida}")
        log(f"📊 Total de registros: {len(df_filtrado)}")
        log("")

        # Mostrar detalle del contenido filtrado
        proveedores = df_filtrado.groupby('Proveedor')
        for prov, grupo in proveedores:
            facturas = ', '.join(grupo['Factura'].astype(str).values)
            montos = grupo['Monto'].values
            monedas = grupo['Moneda'].values
            detalle_montos = ', '.join(f"{m} {c}" for m, c in zip(montos, monedas))
            log(f"  ▸ {prov}")
            log(f"    Facturas: {facturas}")
            log(f"    Montos: {detalle_montos}")

        log("")
        log("✅ Archivo generado correctamente.")
        log("   Ahora puedes usar 'Enviar a Borradores' para crear los correos.")

        # Habilitar botón de enviar
        btn_enviar.configure(state='normal')

    except Exception as e:
        log(f"\n❌ Error en filtrado: {str(e)}")
        messagebox.showerror("Error en filtrado", str(e))


def enviar_a_borradores():
    """PASO 2: Lee un archivo Excel filtrado y crea correos en borradores."""

    # Seleccionar archivo
    archivo_seleccionado = filedialog.askopenfilename(
        title="Seleccionar archivo filtrado",
        initialdir=os.path.dirname(os.path.abspath(__file__)),
        filetypes=[("Archivos Excel", "*.xlsx *.xls")],
        initialfile=archivo_salida
    )

    if not archivo_seleccionado:
        return

    limpiar_log()
    log("=" * 50)
    log("  PASO 2: CREANDO CORREOS EN BORRADORES")
    log("=" * 50)
    log(f"\n📂 Archivo seleccionado: {os.path.basename(archivo_seleccionado)}")

    try:
        df = pd.read_excel(archivo_seleccionado)
        remitentes_df = pd.read_excel(archivo_remitentes)

        df.columns = df.columns.str.strip()
        remitentes_df.columns = remitentes_df.columns.str.strip().str.upper()

        df[columna_proveedor] = df[columna_proveedor].astype(str).str.strip().str.upper()
        remitentes_df['PROVEEDOR'] = remitentes_df['PROVEEDOR'].astype(str).str.strip().str.upper()

        df = df.merge(
            remitentes_df,
            left_on=columna_proveedor,
            right_on='PROVEEDOR',
            how='left'
        )

        faltantes = df[df[columna_para].isna()][columna_proveedor].unique()
        if len(faltantes) > 0:
            lista = '\n'.join(f' - {f}' for f in faltantes)
            log(f"\n❌ Faltan correos para:\n{lista}")
            messagebox.showerror("Proveedores sin correo", f"Faltan correos para:\n{lista}")
            return

        outlook = win32.Dispatch('Outlook.Application')

        if not os.path.exists(ruta_adjuntos):
            log(f"\n❌ No existe la carpeta de adjuntos:\n{ruta_adjuntos}")
            messagebox.showerror("Error", f"No existe la carpeta:\n{ruta_adjuntos}")
            return

        archivos_pdf = [
            f for f in os.listdir(ruta_adjuntos)
            if f.lower().endswith('.pdf') and not f.startswith('~')
        ]

        correos_creados = 0
        log("")

        for (proveedor, fecha_pago), grupo in df.groupby([columna_proveedor, columna_fecha]):

            correo_para = grupo[columna_para].iloc[0]
            correo_cc = grupo[columna_cc].iloc[0] if columna_cc in grupo.columns else ''

            fecha_pdf = pd.to_datetime(fecha_pago).strftime('%Y%m%d')
            proveedor_key = proveedor.replace(' ', '').upper()

            adjunto_encontrado = None
            for archivo in archivos_pdf:
                archivo_key = archivo.replace(' ', '').upper()
                if archivo_key.startswith(fecha_pdf) and proveedor_key in archivo_key:
                    adjunto_encontrado = os.path.join(ruta_adjuntos, archivo)
                    break

            columnas_excluir = ['PROVEEDOR', columna_para, columna_cc]
            columnas_excluir = [c for c in columnas_excluir if c in grupo.columns]

            tabla_html = grupo.drop(columns=columnas_excluir).to_html(
                index=False, border=0, justify='center'
            )

            estilo_tabla = """
            <style>
            table {border-collapse:collapse;width:100%;font-family:Arial;font-size:13px;}
            th {background-color:#1f3a5f;color:white;padding:6px;border:1px solid #1f3a5f;text-align:center;}
            td {padding:6px;border:1px solid #1f3a5f;text-align:center;}
            </style>
            """

            cuerpo_html = f"""
            <html>
            <head>{estilo_tabla}</head>
            <body style="font-family:Arial;font-size:13px;">
            <p>Buenos días,</p>
            <p>Adjuntamos comprobante de pago correspondiente a las siguientes facturas:</p>
            {tabla_html}
            <br>
            <p>Ante cualquier duda, quedo a disposición.</p>
            <p>Buen resto de día,</p>
            <p>Saludos cordiales,</p>
            <p>
            <b>VICTOR DOMINGUEZ</b><br>
            Nutrex INC<br>
            tesoreria2@nutrexinc.com
            </p>
            </body>
            </html>
            """

            mail = outlook.CreateItem(0)
            mail.To = correo_para
            if correo_cc:
                mail.CC = correo_cc
            mail.Subject = f"COMPROBANTE DE PAGO {pd.to_datetime(fecha_pago).strftime('%d/%m/%Y')} - {proveedor}"
            mail.HTMLBody = cuerpo_html

            log(f"📧 Correo creado → {proveedor}")
            log(f"   Para: {correo_para}")
            if correo_cc:
                log(f"   CC: {correo_cc}")

            if adjunto_encontrado:
                mail.Attachments.Add(adjunto_encontrado)
                log(f"   📎 Adjunto: {os.path.basename(adjunto_encontrado)}")
            else:
                log(f"   ⚠️ SIN ADJUNTO (PDF no encontrado)")

            mail.Save()
            correos_creados += 1
            log("")

        log("=" * 50)
        log(f"✅ Se crearon {correos_creados} correo(s) en borradores de Outlook.")
        log("=" * 50)
        messagebox.showinfo("Listo ✅", f"Se crearon {correos_creados} correo(s) en borradores de Outlook.")

    except Exception as e:
        log(f"\n❌ Error al crear correos: {str(e)}")
        messagebox.showerror("Error al crear correos", str(e))


# =====================================================
# === VENTANA TKINTER ===
# =====================================================

if __name__ == '__main__':
    # Cambiar al directorio del script para que los archivos relativos funcionen
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    ventana = tk.Tk()
    ventana.title("Correos Tesorería - Nutrex INC")
    ventana.geometry("650x580")
    ventana.resizable(False, False)
    ventana.configure(bg="#f0f0f0")

    # Centrar ventana en pantalla
    ventana.update_idletasks()
    x = (ventana.winfo_screenwidth() // 2) - (650 // 2)
    y = (ventana.winfo_screenheight() // 2) - (580 // 2)
    ventana.geometry(f"+{x}+{y}")

    tk.Label(
        ventana, text="📧 Generador de Correos - Tesorería",
        font=("Arial", 14, "bold"), bg="#f0f0f0"
    ).pack(pady=(15, 5))

    tk.Label(
        ventana,
        text="Ingresa fecha(s) de pago (YYYY-MM-DD), separadas por coma.\nDeja vacío para usar la fecha más reciente.",
        font=("Arial", 10), bg="#f0f0f0", justify="center"
    ).pack(pady=(0, 5))

    entrada_fecha = tk.Entry(ventana, font=("Arial", 12), width=35, justify="center")
    entrada_fecha.pack(pady=5)
    entrada_fecha.insert(0, datetime.today().strftime('%Y-%m-%d'))

    # Frame para los dos botones
    frame_botones = tk.Frame(ventana, bg="#f0f0f0")
    frame_botones.pack(pady=10)

    tk.Button(
        frame_botones, text="1. Generar Archivo",
        font=("Arial", 11, "bold"), bg="#1f3a5f", fg="white",
        activebackground="#2d5280", activeforeground="white",
        cursor="hand2", width=20,
        command=generar_archivo
    ).pack(side=tk.LEFT, padx=10)

    btn_enviar = tk.Button(
        frame_botones, text="2. Enviar a Borradores",
        font=("Arial", 11, "bold"), bg="#2d7d2d", fg="white",
        activebackground="#3a9a3a", activeforeground="white",
        cursor="hand2", width=20,
        command=enviar_a_borradores, state='disabled'
    )
    btn_enviar.pack(side=tk.LEFT, padx=10)

    # Área de comentarios/log
    tk.Label(
        ventana, text="Comentarios:",
        font=("Arial", 10, "bold"), bg="#f0f0f0", anchor="w"
    ).pack(fill='x', padx=15, pady=(10, 2))

    txt_log = scrolledtext.ScrolledText(
        ventana, font=("Consolas", 9), width=75, height=18,
        state='disabled', bg="#1e1e1e", fg="#d4d4d4",
        insertbackground="white", wrap=tk.WORD
    )
    txt_log.pack(padx=15, pady=(0, 15))

    ventana.mainloop()
