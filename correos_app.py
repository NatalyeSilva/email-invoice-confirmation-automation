import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import pandas as pd
import win32com.client as win32
import os
import threading
from datetime import date


# === CONFIGURACIÓN DE COLUMNAS POR COMBINACIÓN ===

CONFIGURACIONES = {
    ("ANTICIPO", "SECO"): {
        "hoja": "Detalles (Seco)",
        "columna_fecha": "FECHA PAGO ANTICIPO",
        "columnas": [
            'TRANSPORTADORA', 'PLACA', 'CHOFER', 'LICENCIA', 'FACTURA', 'BOOKING',
            'PESO PROGR. TM', 'PRODUCTO', 'FLETE NEGOCIADO2', 'FLETE NEGOCIADO',
            'MONTO ANTICIPO', 'FECHA PAGO ANTICIPO', 'BANCO ANTICIPO', 'SALDO FLETE A CANCELAR',
        ],
        "archivo_salida": "datos_filtrados_anticipo_seco.xlsx",
        "texto_correo": "En adjunto el comprobante y el detalle de pago de Anticipo Seco realizado.",
        "subject_tipo": "ANTICIPO",
    },
    ("SALDO", "SECO"): {
        "hoja": "Detalles (Seco)",
        "columna_fecha": "FECHA PAGO SALDO FLETE",
        "columnas": [
            'TRANSPORTADORA', 'PLACA', 'CHOFER', 'LICENCIA', 'FACTURA', 'BOOKING',
            'PESO PROGR. TM', 'PESO NETO DESCARGADO', 'PRODUCTO', 'FLETE NEGOCIADO2',
            'FLETE NEGOCIADO', 'SALDO FLETE A CANCELAR', 'FECHA PAGO SALDO FLETE', 'BANCO SALDO',
        ],
        "archivo_salida": "datos_filtrados_saldo_seco.xlsx",
        "texto_correo": "En adjunto el comprobante y el detalle de pago de saldo realizado.",
        "subject_tipo": "SALDO",
    },
    ("ANTICIPO", "LIQUIDOS"): {
        "hoja": "Detalles (Liquido)",
        "columna_fecha": "FECHA PAGO ANTICIPO",
        "columnas": [
            'TRANSPORTADORA', 'PLACA', 'CHOFER', 'LICENCIA', 'PESO PROGRAMADO (TM)',
            'PRODUCTO', 'FACTURA', 'BOOKING;FACTURA NUTREX', 'FLETE NEGOCIADO',
            'ANTICIPO', 'FECHA PAGO ANTICIPO', 'BANCO ANTICIPO', 'OBS',
        ],
        "archivo_salida": "datos_filtrados_anticipo_liquido.xlsx",
        "texto_correo": "En adjunto el comprobante y el detalle de pago de anticipo (aceite) realizado.",
        "subject_tipo": "ANTICIPO LIQUID.",
    },
    ("SALDO", "LIQUIDOS"): {
        "hoja": "Detalles (Liquido)",
        "columna_fecha": "FECHA SALDO",
        "columnas": [
            'TRANSPORTADORA', 'PLACA', 'CHOFER', 'LICENCIA', 'PESO PROGRAMADO (TM)',
            'PESO NETO DESCARGADO', 'PRODUCTO', 'FACTURA', 'BOOKING;FACTURA NUTREX',
            'FLETE NEGOCIADO', 'SALDO', 'FECHA SALDO', 'BANCO SALDO', 'OBS',
        ],
        "archivo_salida": "datos_filtrados_saldo_liquido.xlsx",
        "texto_correo": "En adjunto el comprobante y el detalle de pago de saldo líquido realizado.",
        "subject_tipo": "SALDO LIQUID.",
    },
}

RUTA_ADJUNTOS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "COMPROBANTES")
ARCHIVO_PLANTILLA = "Plantilla.xlsx"
ARCHIVO_REMITENTES = "remitentes.xlsx"

ESTILO_TABLA = """
<style>
    table {
        border-collapse: collapse;
        width: 100%;
        font-family: Arial, sans-serif;
        font-size: 12px;
    }
    th {
        background-color: #4A90E2;
        color: white;
        font-weight: bold;
        padding: 8px;
        border: 1px solid #ddd;
        text-align: center;
    }
    td {
        padding: 8px;
        border: 1px solid #ddd;
        text-align: center;
    }
    tr:nth-child(even) {
        background-color: #f2f2f2;
    }
</style>
"""


class CorreosApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Correos Transportistas - NUTREX")
        self.root.geometry("700x550")
        self.root.resizable(False, False)

        # Variables
        self.modo_var = tk.StringVar(value="ANTICIPO")
        self.producto_var = tk.StringVar(value="SECO")
        self.archivo_salida = None
        self.fecha_filtrada = None

        self._crear_interfaz()
        self._verificar_planilla_existente()

    def _crear_interfaz(self):
        # === Frame superior: selección ===
        frame_top = ttk.LabelFrame(self.root, text="Configuración", padding=15)
        frame_top.pack(fill="x", padx=15, pady=(15, 5))

        # Modo
        ttk.Label(frame_top, text="Modo:", font=("Arial", 11, "bold")).grid(row=0, column=0, padx=5, pady=8, sticky="w")
        combo_modo = ttk.Combobox(frame_top, textvariable=self.modo_var, values=["ANTICIPO", "SALDO"],
                                  state="readonly", width=18, font=("Arial", 11))
        combo_modo.grid(row=0, column=1, padx=5, pady=8)
        combo_modo.bind("<<ComboboxSelected>>", lambda e: self._verificar_planilla_existente())

        # Producto
        ttk.Label(frame_top, text="Producto:", font=("Arial", 11, "bold")).grid(row=0, column=2, padx=5, pady=8, sticky="w")
        combo_prod = ttk.Combobox(frame_top, textvariable=self.producto_var, values=["SECO", "LIQUIDOS"],
                                  state="readonly", width=18, font=("Arial", 11))
        combo_prod.grid(row=0, column=3, padx=5, pady=8)
        combo_prod.bind("<<ComboboxSelected>>", lambda e: self._verificar_planilla_existente())

        # === Frame botones ===
        frame_btn = ttk.Frame(self.root, padding=5)
        frame_btn.pack(fill="x", padx=15, pady=5)

        self.btn_generar = ttk.Button(frame_btn, text="1. Generar Planilla", command=self._on_generar)
        self.btn_generar.pack(side="left", padx=5, expand=True, fill="x")

        self.btn_borradores = ttk.Button(frame_btn, text="2. Enviar a Borradores", command=self._on_borradores, state="disabled")
        self.btn_borradores.pack(side="left", padx=5, expand=True, fill="x")

        # === Frame log ===
        frame_log = ttk.LabelFrame(self.root, text="Estado", padding=10)
        frame_log.pack(fill="both", expand=True, padx=15, pady=(5, 15))

        self.log = scrolledtext.ScrolledText(frame_log, height=18, font=("Consolas", 10), state="disabled", wrap="word")
        self.log.pack(fill="both", expand=True)

    # --- Verificar planilla existente ---
    def _verificar_planilla_existente(self):
        config = self._get_config()
        if not config:
            return
        archivo = config["archivo_salida"]
        if os.path.exists(archivo):
            try:
                df_check = pd.read_excel(archivo)
                candidatas = ['FECHA PAGO ANTICIPO', 'FECHA PAGO SALDO FLETE', 'FECHA SALDO']
                col_fecha = next((c for c in candidatas if c in df_check.columns), None)
                if col_fecha:
                    fecha = pd.to_datetime(df_check[col_fecha], errors='coerce').max()
                else:
                    fecha = pd.Timestamp(date.today())
                self.archivo_salida = archivo
                self.fecha_filtrada = fecha.date() if hasattr(fecha, 'date') else fecha
                self.btn_borradores.config(state="normal")
                self._log_clear()
                self._log(f"📂 Planilla existente detectada: {archivo}")
                self._log(f"   Fecha: {self.fecha_filtrada} | Filas: {len(df_check)}")
                self._log(f"   Podés generar de nuevo o enviar a borradores directamente.")
            except Exception:
                self.archivo_salida = None
                self.btn_borradores.config(state="disabled")
        else:
            self.archivo_salida = None
            self.fecha_filtrada = None
            self.btn_borradores.config(state="disabled")
            self._log_clear()
            self._log(f"ℹ️ No existe planilla para {self.modo_var.get()} - {self.producto_var.get()}. Generá una primero.")

    # --- Logging ---
    def _log(self, msg):
        self.log.config(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.config(state="disabled")

    def _log_clear(self):
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")

    # --- Obtener config ---
    def _get_config(self):
        modo = self.modo_var.get()
        producto = self.producto_var.get()
        key = (modo, producto)
        if key not in CONFIGURACIONES:
            return None
        return CONFIGURACIONES[key]

    # --- Paso 1: Generar planilla ---
    def _on_generar(self):
        self._log_clear()
        self.btn_generar.config(state="disabled")
        self.btn_borradores.config(state="disabled")
        threading.Thread(target=self._generar_planilla, daemon=True).start()

    def _generar_planilla(self):
        try:
            config = self._get_config()
            modo = self.modo_var.get()
            producto = self.producto_var.get()

            self._log(f"=== GENERANDO PLANILLA: {modo} - {producto} ===")
            self._log(f"Hoja: {config['hoja']}")
            self._log(f"Columna fecha: {config['columna_fecha']}")

            df = pd.read_excel(ARCHIVO_PLANTILLA, sheet_name=config["hoja"])
            col_fecha = config["columna_fecha"]
            df[col_fecha] = pd.to_datetime(df[col_fecha], errors="coerce")
            fecha_max = df[col_fecha].max().date()
            df_filtrado = df[df[col_fecha].dt.date == fecha_max]

            faltantes = [c for c in config["columnas"] if c not in df_filtrado.columns]
            if faltantes:
                self._log(f"❌ Columnas faltantes: {faltantes}")
                self.root.after(0, lambda: self.btn_generar.config(state="normal"))
                return

            df_filtrado = df_filtrado[config["columnas"]]
            archivo = config["archivo_salida"]
            df_filtrado.to_excel(archivo, index=False)

            self.archivo_salida = archivo
            self.fecha_filtrada = fecha_max

            self._log(f"\n✅ Archivo generado: {archivo}")
            self._log(f"Fecha filtrada: {fecha_max}")
            self._log(f"Filas: {len(df_filtrado)}")

            self.root.after(0, lambda: self.btn_generar.config(state="normal"))
            self.root.after(0, lambda: self.btn_borradores.config(state="normal"))

        except Exception as e:
            self._log(f"\n❌ ERROR: {e}")
            self.root.after(0, lambda: self.btn_generar.config(state="normal"))

    # --- Paso 2: Enviar a borradores ---
    def _on_borradores(self):
        if not self.archivo_salida:
            messagebox.showwarning("Aviso", "Primero generá la planilla.")
            return

        respuesta = messagebox.askyesno("Confirmar",
            f"¿Enviar correos a Borradores?\n\n"
            f"Modo: {self.modo_var.get()}\n"
            f"Producto: {self.producto_var.get()}\n"
            f"Archivo: {self.archivo_salida}\n"
            f"Fecha: {self.fecha_filtrada}")

        if not respuesta:
            return

        self.btn_generar.config(state="disabled")
        self.btn_borradores.config(state="disabled")
        threading.Thread(target=self._enviar_borradores, daemon=True).start()

    def _enviar_borradores(self):
        try:
            config = self._get_config()
            self._log(f"\n=== ENVIANDO A BORRADORES ===")

            df = pd.read_excel(self.archivo_salida)
            remitentes_df = pd.read_excel(ARCHIVO_REMITENTES)

            col_trans = "TRANSPORTADORA"
            col_para = "PARA"
            col_cc = "CC"

            df[col_trans] = df[col_trans].str.strip().str.upper()
            remitentes_df[col_trans] = remitentes_df[col_trans].str.strip().str.upper()

            df = df.merge(remitentes_df, on=col_trans, how="left")

            faltantes = df[df[col_para].isna()][col_trans].unique()
            if len(faltantes) > 0:
                self._log("❌ Transportadoras sin correo en 'PARA':")
                for f in faltantes:
                    self._log(f"  - {f}")
                self.root.after(0, lambda: self.btn_generar.config(state="normal"))
                self.root.after(0, lambda: self.btn_borradores.config(state="normal"))
                return

            # Determinar fecha
            candidatas = ['FECHA PAGO ANTICIPO', 'FECHA PAGO SALDO FLETE', 'FECHA SALDO']
            col_fecha = next((c for c in candidatas if c in df.columns), None)
            if col_fecha:
                df[col_fecha] = pd.to_datetime(df[col_fecha], errors='coerce')
                fecha_mas_reciente = df[col_fecha].max()
            else:
                posibles = [c for c in df.columns if 'FECHA' in c.upper()]
                if posibles:
                    fecha_mas_reciente = pd.to_datetime(df[posibles[0]], errors='coerce').max()
                else:
                    fecha_mas_reciente = pd.Timestamp(date.today())

            outlook = win32.Dispatch('Outlook.Application')
            subject_tipo = config["subject_tipo"]
            texto_correo = config["texto_correo"]
            count = 0

            for transportadora, grupo in df.groupby(col_trans):
                correo_para = grupo[col_para].iloc[0]
                correo_cc = grupo[col_cc].iloc[0] if col_cc in grupo.columns and pd.notna(grupo[col_cc].iloc[0]) else ''

                nombre_busqueda = transportadora.upper()

                # Buscar adjunto PDF
                adjunto_encontrado = None
                if os.path.isdir(RUTA_ADJUNTOS):
                    for archivo in os.listdir(RUTA_ADJUNTOS):
                        if archivo.lower().endswith('.pdf') and nombre_busqueda in archivo.upper():
                            adjunto_encontrado = os.path.join(RUTA_ADJUNTOS, archivo)
                            break

                # Tabla HTML
                cols_drop = [c for c in [col_para, col_cc] if c in grupo.columns]
                tabla_html = grupo.drop(columns=cols_drop).to_html(index=False, border=0, justify='center', escape=False)

                cuerpo_html = f"""
                <html>
                <head>{ESTILO_TABLA}</head>
                <body>
                    <p>Buenos días,</p>
                    <p>{texto_correo}</p>

                    {tabla_html}

                    <p>Por favor si pueden enviar un recibo como comprobante de que el pago fue recibido, ¡muchas gracias!</p>
                    <p>Tener en cuenta que no recibiremos reclamos posteriores a los dos meses de emisión de este comprobante.</p>
                    <p>Saludos cordiales,<br></p>

                    <span>Natalye Silva - Verificaciones,</span><br>
                    <span>NUTREX INC.</span><br>
                    <span>verificacion@nutrexinc.com</span>
                </body>
                </html>
                """

                mail = outlook.CreateItem(0)
                mail.To = correo_para
                if correo_cc:
                    mail.CC = correo_cc
                mail.Subject = f"PAGO TRANSPORTISTA {fecha_mas_reciente.strftime('%d/%m/%Y')} - {subject_tipo} - {transportadora}"
                mail.HTMLBody = cuerpo_html

                if adjunto_encontrado and os.path.exists(adjunto_encontrado):
                    mail.Attachments.Add(adjunto_encontrado)
                    self._log(f"📎 Adjunto: {transportadora} → {os.path.basename(adjunto_encontrado)}")
                else:
                    self._log(f"⚠️ Sin adjunto para {transportadora}")

                mail.Save()
                count += 1

            self._log(f"\n✅ {count} correos guardados en Borradores.")
            self.root.after(0, lambda: self.btn_generar.config(state="normal"))
            self.root.after(0, lambda: self.btn_borradores.config(state="normal"))

        except Exception as e:
            self._log(f"\n❌ ERROR: {e}")
            self.root.after(0, lambda: self.btn_generar.config(state="normal"))
            self.root.after(0, lambda: self.btn_borradores.config(state="normal"))


if __name__ == "__main__":
    root = tk.Tk()
    app = CorreosApp(root)
    root.mainloop()
