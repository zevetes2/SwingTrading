import gspread
from oauth2client.service_account import ServiceAccountCredentials
import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime
import time
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.gridspec as gridspec
from matplotlib.patches import FancyArrowPatch
import warnings
warnings.filterwarnings('ignore')

# ─── Funciones auxiliares de columnas ────────────────────────────────────────
def col_to_letter(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def letter_to_col(label):
    res = 0
    for char in label:
        res = res * 26 + (ord(char.upper()) - ord('A') + 1)
    return res


# ─── Sistema Principal ────────────────────────────────────────────────────────
class WilliamsRSwingSystem:
    def __init__(self, credenciales_json='principios.json',
                 nombre_archivo="Portafolio Tracker",
                 nombre_hoja="Swing Trading"):
        print("🔄 Conectando con Google Sheets...")
        try:
            scope = ['https://spreadsheets.google.com/feeds',
                     'https://www.googleapis.com/auth/drive']
            creds = ServiceAccountCredentials.from_json_keyfile_name(credenciales_json, scope)
            self.client = gspread.authorize(creds)
            self.workbook = self.client.open(nombre_archivo)
            self.sheet = self.workbook.worksheet(nombre_hoja)
            print(f"✅ Conectado a: {nombre_archivo} -> {nombre_hoja}\n")
        except Exception as e:
            print(f"❌ Error de conexión: {e}")
            raise

    # ── Indicadores ──────────────────────────────────────────────────────────
    def calcular_williams_r(self, df, period=14):
        highest_high = df['High'].rolling(window=period).max()
        lowest_low   = df['Low'].rolling(window=period).min()
        return ((highest_high - df['Close']) / (highest_high - lowest_low)) * -100

    def detectar_divergencias(self, df, williams_r, lookback=10):
        precio = df['Close'].values
        wr     = williams_r.values
        if len(precio) < lookback:
            return "Ninguna"
        p_reciente  = precio[-lookback:]
        wr_reciente = wr[-lookback:]
        if (p_reciente[-1] < np.min(p_reciente[:-1]) and
                wr_reciente[-1] > np.min(wr_reciente[:-1]) and
                wr_reciente[-1] < -50):
            return "🟢 Alcista (Bullish)"
        if (p_reciente[-1] > np.max(p_reciente[:-1]) and
                wr_reciente[-1] < np.max(wr_reciente[:-1]) and
                wr_reciente[-1] > -50):
            return "🔴 Bajista (Bearish)"
        return "Ninguna"

    # ── Análisis completo de un ticker ───────────────────────────────────────
    def analizar_ticker_completo(self, ticker):
        try:
            print(f"  📊 Analizando {ticker}...", end=" ", flush=True)
            stock = yf.Ticker(ticker)
            df_w  = stock.history(period="2y", interval="1wk")
            if df_w.empty or len(df_w) < 30:
                print("❌ Datos insuficientes")
                return None

            w14   = self.calcular_williams_r(df_w, 14)
            w7    = self.calcular_williams_r(df_w, 7)
            w28   = self.calcular_williams_r(df_w, 28)
            wr_act = w14.iloc[-1]
            precio = df_w['Close'].iloc[-1]
            sma20  = df_w['Close'].rolling(20).mean().iloc[-1]
            sma50  = df_w['Close'].rolling(50).mean().iloc[-1] if len(df_w) >= 50 else sma20
            v_ratio = (df_w['Volume'].iloc[-1] /
                       df_w['Volume'].rolling(10).mean().iloc[-1])
            tr  = pd.concat([
                df_w['High'] - df_w['Low'],
                abs(df_w['High'] - df_w['Close'].shift()),
                abs(df_w['Low']  - df_w['Close'].shift())
            ], axis=1).max(axis=1)
            atr = tr.rolling(14).mean().iloc[-1]

            puntos = 50
            if wr_act <= -80:  puntos += 25
            elif wr_act >= -20: puntos -= 25
            div = self.detectar_divergencias(df_w, w14)
            if "Alcista" in div: puntos += 20
            if precio > sma20:   puntos += 10

            if   puntos >= 80: dec, acc = "🟢🟢 COMPRA FUERTE", "COMPRAR"
            elif puntos >= 65: dec, acc = "🟢 COMPRA",          "COMPRAR"
            elif puntos >= 40: dec, acc = "🟡 OBSERVAR",        "VIGILAR"
            else:              dec, acc = "🔴 EVITAR",           "ESPERAR"

            sl  = round(precio - (atr * 2), 2)
            tp1 = round(precio + (atr * 3), 2)
            tp2 = round(precio + (atr * 5), 2)
            rb  = round(((tp1 - precio) / (precio - sl)), 2) if precio != sl else 0

            print(f"Puntos: {puntos}")
            return {
                'ticker': ticker,
                'precio': round(precio, 2),
                'r14': round(wr_act, 1), 'r7': round(w7.iloc[-1], 1),
                'r28': round(w28.iloc[-1], 1),
                'zona': ("Sobreventa" if wr_act <= -80
                         else "Sobrecompra" if wr_act >= -20 else "Neutral"),
                'div': div,
                'sma10': round(df_w['Close'].rolling(10).mean().iloc[-1], 2),
                'sma20': round(sma20, 2), 'sma50': round(sma50, 2),
                'vol': round(v_ratio, 2), 'atr': round(atr, 2),
                'pts': puntos, 'dec': dec, 'acc': acc,
                'rec': ("Setup Perfecto"
                        if (wr_act <= -80 and "Alcista" in div)
                        else "Análisis Normal"),
                'sig': f"WR: {round(wr_act,1)} | Vol: {round(v_ratio,1)}x",
                'sl': sl, 'tp1': tp1, 'tp2': tp2, 'rb': rb,
                'df_w': df_w, 'w14': w14, 'w7': w7
            }
        except Exception as e:
            print(f"❌ Error: {e}")
            return None

    # ── Gráfica de Swing para un ticker ─────────────────────────────────────
    def generar_grafica_swing(self, datos, guardar_como=None):
        """
        Genera una gráfica profesional de 3 paneles:
          1. Precio + SMA + Zona de entrada + SL/TP
          2. Williams %R (14, 7, 28)
          3. Volumen relativo
        """
        ticker = datos['ticker']
        df     = datos['df_w'].copy().iloc[-52:]   # último año semanal
        w14    = datos['w14'].iloc[-52:]
        w7     = datos['w7'].iloc[-52:]
        precio = datos['precio']
        sl     = datos['sl']
        tp1    = datos['tp1']
        tp2    = datos['tp2']
        sma20  = datos['sma20']
        sma50  = datos['sma50']
        puntos = datos['pts']
        dec    = datos['dec']
        div    = datos['div']

        # Paleta oscura estilo trading
        BG      = '#0d1117'
        PANEL   = '#161b22'
        GREEN   = '#26a641'
        RED     = '#f85149'
        YELLOW  = '#e3b341'
        BLUE    = '#58a6ff'
        PURPLE  = '#bc8cff'
        ORANGE  = '#f0883e'
        GRAY    = '#8b949e'
        WHITE   = '#e6edf3'

        fig = plt.figure(figsize=(16, 12), facecolor=BG)
        gs  = gridspec.GridSpec(3, 1, figure=fig,
                                height_ratios=[3, 1.5, 1],
                                hspace=0.08)

        ax1 = fig.add_subplot(gs[0])   # Precio
        ax2 = fig.add_subplot(gs[1], sharex=ax1)  # Williams %R
        ax3 = fig.add_subplot(gs[2], sharex=ax1)  # Volumen

        for ax in [ax1, ax2, ax3]:
            ax.set_facecolor(PANEL)
            ax.tick_params(colors=GRAY, labelsize=8)
            ax.spines[:].set_color('#30363d')

        fechas = df.index

        # ── Panel 1: Precio ──────────────────────────────────────────────────
        # Velas simplificadas (líneas High-Low + rect Open-Close)
        for i, (idx, row) in enumerate(df.iterrows()):
            color = GREEN if row['Close'] >= row['Open'] else RED
            ax1.plot([i, i], [row['Low'], row['High']], color=color, lw=0.8, alpha=0.7)
            rect = plt.Rectangle((i - 0.3, min(row['Open'], row['Close'])),
                                  0.6, abs(row['Close'] - row['Open']),
                                  color=color, alpha=0.85)
            ax1.add_patch(rect)

        n = len(df)
        xs = range(n)

        # SMAs
        sma20_s = df['Close'].rolling(20).mean()
        sma50_s = df['Close'].rolling(50).mean()
        ax1.plot(xs, sma20_s.values, color=BLUE,   lw=1.2, label='SMA 20', alpha=0.9)
        ax1.plot(xs, sma50_s.values, color=PURPLE, lw=1.2, label='SMA 50', alpha=0.9, ls='--')

        # Zona de entrada (último precio actual)
        ax1.axhline(precio, color=YELLOW, lw=1.5, ls='-',  alpha=0.9, label=f'Entrada: ${precio}')
        ax1.axhline(sl,     color=RED,    lw=1.5, ls='--', alpha=0.9, label=f'Stop Loss: ${sl}')
        ax1.axhline(tp1,    color=GREEN,  lw=1.5, ls='--', alpha=0.9, label=f'TP1: ${tp1}')
        ax1.axhline(tp2,    color=GREEN,  lw=2.0, ls='-',  alpha=0.9, label=f'TP2: ${tp2}')

        # Rellenos de zonas
        ax1.axhspan(sl, precio, alpha=0.08, color=RED,   label='_zona_riesgo')
        ax1.axhspan(precio, tp1, alpha=0.06, color=GREEN, label='_zona_tp1')
        ax1.axhspan(tp1, tp2,   alpha=0.10, color=GREEN, label='_zona_tp2')

        # Flechas de SL y TPs en el margen derecho
        x_ann = n + 0.5
        ax1.annotate(f'TP2 ${tp2}', xy=(n-1, tp2), xytext=(n+1, tp2),
                     color=GREEN, fontsize=8, fontweight='bold', va='center',
                     arrowprops=dict(arrowstyle='->', color=GREEN, lw=1))
        ax1.annotate(f'TP1 ${tp1}', xy=(n-1, tp1), xytext=(n+1, tp1),
                     color=GREEN, fontsize=8, va='center',
                     arrowprops=dict(arrowstyle='->', color=GREEN, lw=1))
        ax1.annotate(f'ENTRADA ${precio}', xy=(n-1, precio), xytext=(n+1, precio),
                     color=YELLOW, fontsize=8, fontweight='bold', va='center',
                     arrowprops=dict(arrowstyle='->', color=YELLOW, lw=1))
        ax1.annotate(f'SL ${sl}', xy=(n-1, sl), xytext=(n+1, sl),
                     color=RED, fontsize=8, fontweight='bold', va='center',
                     arrowprops=dict(arrowstyle='->', color=RED, lw=1.2))

        ax1.set_xlim(-1, n + 10)
        ax1.legend(loc='upper left', fontsize=7.5, facecolor=PANEL,
                   labelcolor=WHITE, framealpha=0.8, ncol=3)
        ax1.set_ylabel('Precio (USD)', color=GRAY, fontsize=9)
        ax1.tick_params(labelbottom=False)

        # Info de R:R
        rb = datos['rb']
        rb_color = GREEN if rb >= 2 else YELLOW if rb >= 1.5 else RED
        ax1.text(0.01, 0.97,
                 f"R:R = {rb}x  |  Riesgo: ${round(precio-sl,2)}  |  Ganancia TP1: ${round(tp1-precio,2)}",
                 transform=ax1.transAxes, color=rb_color,
                 fontsize=8.5, va='top', fontweight='bold',
                 bbox=dict(boxstyle='round,pad=0.3', facecolor=BG, alpha=0.7))

        # ── Panel 2: Williams %R ─────────────────────────────────────────────
        ax2.plot(xs, w14.values, color=BLUE,   lw=1.5, label='W%R 14', alpha=0.95)
        ax2.plot(xs, w7.values,  color=ORANGE, lw=1.0, label='W%R 7',  alpha=0.7, ls='--')

        ax2.axhline(-20, color=RED,   lw=0.8, ls=':', alpha=0.7)
        ax2.axhline(-50, color=GRAY,  lw=0.8, ls=':', alpha=0.5)
        ax2.axhline(-80, color=GREEN, lw=0.8, ls=':', alpha=0.7)

        ax2.axhspan(-100, -80, alpha=0.12, color=GREEN)
        ax2.axhspan(-20,    0, alpha=0.12, color=RED)

        ax2.text(0, -15, 'Sobrecompra', color=RED,   fontsize=7, alpha=0.8)
        ax2.text(0, -95, 'Sobreventa',  color=GREEN, fontsize=7, alpha=0.8)

        ax2.set_ylim(-105, 5)
        ax2.set_ylabel('Williams %R', color=GRAY, fontsize=9)
        ax2.legend(loc='upper right', fontsize=7.5, facecolor=PANEL,
                   labelcolor=WHITE, framealpha=0.8)
        ax2.tick_params(labelbottom=False)

        # Marcar el valor actual de WR
        wr_actual = w14.iloc[-1]
        wr_color  = GREEN if wr_actual <= -80 else RED if wr_actual >= -20 else YELLOW
        ax2.scatter([n-1], [wr_actual], color=wr_color, s=60, zorder=5)
        ax2.text(n-1, wr_actual + 3, f'{wr_actual:.1f}',
                 color=wr_color, fontsize=7.5, ha='center', fontweight='bold')

        # ── Panel 3: Volumen relativo ────────────────────────────────────────
        vol_med = df['Volume'].rolling(10).mean()
        for i, (v, vm) in enumerate(zip(df['Volume'].values, vol_med.values)):
            color = GREEN if v > vm else GRAY
            ax3.bar(i, v, color=color, alpha=0.75, width=0.8)
        ax3.plot(xs, vol_med.values, color=YELLOW, lw=1.0, ls='--', label='Vol MA10')
        ax3.set_ylabel('Volumen', color=GRAY, fontsize=9)
        ax3.legend(loc='upper right', fontsize=7.5, facecolor=PANEL,
                   labelcolor=WHITE, framealpha=0.8)

        # ── Etiquetas del eje X ──────────────────────────────────────────────
        step = max(1, n // 8)
        xticks = list(range(0, n, step))
        xlabels = [fechas[i].strftime('%b %Y') for i in xticks]
        ax3.set_xticks(xticks)
        ax3.set_xticklabels(xlabels, color=GRAY, fontsize=8, rotation=30)

        # ── Título principal ─────────────────────────────────────────────────
        color_titulo = GREEN if 'COMPRA' in dec else YELLOW if 'OBSERVAR' in dec else RED
        fig.suptitle(
            f"{ticker}  —  {dec}  |  Puntuación: {puntos}/100\n"
            f"Divergencia: {div}  |  Zona WR: {datos['zona']}  |  Vol Ratio: {datos['vol']}x",
            color=WHITE, fontsize=13, fontweight='bold', y=0.98,
            bbox=dict(boxstyle='round,pad=0.4', facecolor=color_titulo, alpha=0.25)
        )

        plt.tight_layout(rect=[0, 0, 1, 0.96])

        # Guardar
        nombre = guardar_como or f"swing_{ticker}_{datetime.now().strftime('%Y%m%d_%H%M')}.png"
        plt.savefig(nombre, dpi=150, bbox_inches='tight', facecolor=BG)
        plt.close()
        print(f"      📈 Gráfica guardada: {nombre}")
        return nombre

    # ── Actualizar hoja y generar top-5 ─────────────────────────────────────
    def actualizar_hoja(self, fila_i, fila_f, col_t, col_s, top_n=5):
        col_data = self.sheet.col_values(col_t)
        f_final  = fila_f if fila_f else len(col_data)
        tickers  = [t.strip().upper() for t in col_data[fila_i-1:f_final] if t.strip()]

        if not tickers:
            return print("⚠️  No hay tickers en la columna indicada.")

        # ── Análisis de todos los tickers ───────────────────────────────────
        todos_resultados = []
        filas_sheet      = []

        for t in tickers:
            d = self.analizar_ticker_completo(t)
            if d:
                res = [d['r14'], d['r7'], d['r28'], d['zona'], d['div'],
                       d['sma10'], d['sma20'], d['sma50'], d['vol'], d['atr'],
                       d['pts'], d['dec'], d['acc'], d['rec'], d['sig'],
                       d['sl'], d['tp1'], d['tp2'], d['rb']]
                filas_sheet.append(
                    [(0.0 if isinstance(x, float) and (np.isnan(x) or np.isinf(x)) else x)
                     for x in res]
                )
                todos_resultados.append(d)
            else:
                filas_sheet.append(['ERROR'] * 19)

        # ── Escribir en Google Sheets ────────────────────────────────────────
        if filas_sheet:
            col_start_idx  = letter_to_col(col_s)
            col_end_letter = col_to_letter(col_start_idx + 18)
            rango = f"{col_s}{fila_i}:{col_end_letter}{fila_i + len(filas_sheet) - 1}"
            self.sheet.update(range_name=rango, values=filas_sheet)
            print(f"\n✅ Datos actualizados en {rango}")

        # ── Top N por puntuación → gráficas ─────────────────────────────────
        if todos_resultados:
            top = sorted(todos_resultados, key=lambda x: x['pts'], reverse=True)[:top_n]
            print(f"\n{'='*60}")
            print(f"  🏆  TOP {top_n} MEJORES SETUPS — GENERANDO GRÁFICAS")
            print(f"{'='*60}")
            for rank, d in enumerate(top, 1):
                print(f"\n  #{rank}  {d['ticker']}  —  {d['pts']} pts  —  {d['dec']}")
                nombre_archivo = (
                    f"TOP{rank}_{d['ticker']}_"
                    f"{datetime.now().strftime('%Y%m%d_%H%M')}.png"
                )
                self.generar_grafica_swing(d, guardar_como=nombre_archivo)

            print(f"\n✅ {len(top)} gráficas generadas.")


# ─── Configuración ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    ARCHIVO_NOMBRE = "Portafolio Tracker"
    HOJA_NOMBRE    = "Swing Trading"
    FILA_INICIAL   = 2
    FILA_FINAL     = 182      # None para procesar toda la columna
    COL_TICKERS    = 1      # Columna A
    COL_SALIDA     = 'R'    # Columna donde escribir resultados
    TOP_N          = 10      # Cuántos setups graficar

    try:
        sistema = WilliamsRSwingSystem(
            nombre_archivo=ARCHIVO_NOMBRE,
            nombre_hoja=HOJA_NOMBRE
        )
        sistema.actualizar_hoja(
            FILA_INICIAL, FILA_FINAL,
            COL_TICKERS, COL_SALIDA,
            top_n=TOP_N
        )
    except Exception as e:
        print(f"\n❌ Fallo crítico: {e}")
