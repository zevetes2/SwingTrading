"""
╔══════════════════════════════════════════════════════════════════╗
║          SWING TRADING v3 — SISTEMA MTF + LARGO PLAZO           ║
║                                                                  ║
║  Combina:                                                        ║
║   • Puntuación Swing (técnico multi-timeframe: semanal+diario)  ║
║   • Puntuación Largo Plazo (fundamentals de "7 PRINCIPIOS")     ║
║  → Matriz de decisión conjunta para evitar falsos positivos      ║
╚══════════════════════════════════════════════════════════════════╝
"""

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
import warnings
warnings.filterwarnings('ignore')


# ══════════════════════════════════════════════════════════════════
# CONFIGURACIÓN  ← edita solo esta sección
# ══════════════════════════════════════════════════════════════════
CFG = {
    # Google Sheets
    'credenciales_json':  'principios.json',

    # Archivo/hoja del swing
    'swing_archivo':      'Portafolio Tracker',
    'swing_hoja':         'Swing Trading',
    'swing_fila_ini':     2,
    'swing_fila_fin':     183,       # None = toda la columna
    'swing_col_tickers':  1,         # columna A
    'swing_col_salida':   'R',       # primera col de escritura (26 cols)

    # Archivo/hoja del largo plazo (main.py)
    'lp_archivo':         'Portafolio Financiero',
    'lp_hoja':            '7 PRINCIPIOS',
    'lp_fila_ini':        7,         # start_row en main.py
    'lp_fila_fin':        190,        # end_row   en main.py
    'lp_col_ticker':      'A',       # tickers
    'lp_col_puntuacion':  'CQ',       # columna con la puntuación largo plazo
    'lp_col_earnings':    'V',       # columna Earnings Window (de main.py)

    # Comportamiento
    'top_n':              10,        # cuántos setups graficar
}


# ══════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════
def col_to_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def letter_to_col(label):
    r = 0
    for c in label:
        r = r * 26 + (ord(c.upper()) - 64)
    return r

def safe(v):
    """Convierte NaN/Inf a 0.0 para Google Sheets."""
    if isinstance(v, float) and (np.isnan(v) or np.isinf(v)):
        return 0.0
    return v


# ══════════════════════════════════════════════════════════════════
# MATRIZ DE DECISIÓN CONJUNTA
# ══════════════════════════════════════════════════════════════════
def matriz_decision(pts_swing: int, pts_lp) -> dict:
    """
    Combina puntuación swing (0-100) y largo plazo (cualquier escala).
    Retorna etiqueta de acción, emoji y ajuste de puntos para ranking.

    Largo plazo se normaliza internamente a 3 zonas:
      ALTO   ≥ 70   → empresa de calidad
      MEDIO  50-69  → empresa aceptable
      BAJO   < 50   → empresa débil / sin datos
    """
    # Normalizar largo plazo
    if pts_lp is None:
        zona_lp = "BAJO"
    else:
        try:
            lp = float(str(pts_lp).replace('%','').strip())
        except:
            lp = 0
        zona_lp = "ALTO" if lp >= 70 else "MEDIO" if lp >= 50 else "BAJO"

    # Normalizar swing
    zona_sw = "ALTO" if pts_swing >= 65 else "MEDIO" if pts_swing >= 40 else "BAJO"

    # Tabla de decisión
    tabla = {
        # (zona_lp, zona_sw): (acción, emoji, bonus_pts)
        ("ALTO",  "ALTO"):  ("🟢🟢 COMPRA FUERTE MTF+LP",  "COMPRAR",  +15),
        ("ALTO",  "MEDIO"): ("🟢 COMPRA MODERADA",          "COMPRAR",   +5),
        ("ALTO",  "BAJO"):  ("🟡 ESPERAR TIMING",           "VIGILAR",   -5),
        ("MEDIO", "ALTO"):  ("🟢 COMPRA SELECTIVA",         "COMPRAR",   +5),
        ("MEDIO", "MEDIO"): ("🟡 VIGILAR",                  "VIGILAR",    0),
        ("MEDIO", "BAJO"):  ("🔴 NO OPERAR",                "ESPERAR",  -10),
        ("BAJO",  "ALTO"):  ("⚠️ SOLO SCALP CORTO",        "VIGILAR",  -10),
        ("BAJO",  "MEDIO"): ("🔴 EVITAR",                   "ESPERAR",  -15),
        ("BAJO",  "BAJO"):  ("🔴🔴 EVITAR TOTALMENTE",      "ESPERAR",  -20),
    }
    accion, acc, bonus = tabla[(zona_lp, zona_sw)]
    return {
        'accion':   accion,
        'acc':      acc,
        'bonus':    bonus,
        'zona_lp':  zona_lp,
        'zona_sw':  zona_sw,
    }


# ══════════════════════════════════════════════════════════════════
# SISTEMA PRINCIPAL
# ══════════════════════════════════════════════════════════════════
class SwingSystemV3:

    def __init__(self, cfg: dict):
        self.cfg = cfg
        print("🔄 Conectando con Google Sheets...")
        scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            cfg['credenciales_json'], scope)
        self.client = gspread.authorize(creds)

        # Hoja swing
        wb_swing       = self.client.open(cfg['swing_archivo'])
        self.sh_swing  = wb_swing.worksheet(cfg['swing_hoja'])

        # Hoja largo plazo (puede ser otro workbook)
        try:
            if cfg['lp_archivo'] == cfg['swing_archivo']:
                wb_lp = wb_swing
            else:
                wb_lp = self.client.open(cfg['lp_archivo'])
            self.sh_lp = wb_lp.worksheet(cfg['lp_hoja'])
            print(f"✅ Largo plazo: {cfg['lp_archivo']} → {cfg['lp_hoja']}")
        except Exception as e:
            print(f"⚠️  No se pudo conectar a hoja LP: {e}  (se usará N/D)")
            self.sh_lp = None

        print(f"✅ Swing: {cfg['swing_archivo']} → {cfg['swing_hoja']}\n")

    # ──────────────────────────────────────────────────────────────
    # LEER PUNTUACIONES DE LARGO PLAZO
    # ──────────────────────────────────────────────────────────────
    def leer_puntuaciones_lp(self) -> dict:
        """
        Lee tickers y puntuaciones de la hoja '7 PRINCIPIOS'.
        Retorna dict  {TICKER: puntuacion_numerica}
        """
        if self.sh_lp is None:
            return {}
        cfg = self.cfg
        try:
            fi = cfg['lp_fila_ini']
            ff = cfg['lp_fila_fin']
            col_t = cfg['lp_col_ticker']
            col_p = cfg['lp_col_puntuacion']

            col_e = cfg.get('lp_col_earnings', 'V')
            tickers_raw  = self.sh_lp.get(f"{col_t}{fi}:{col_t}{ff}")
            puntos_raw   = self.sh_lp.get(f"{col_p}{fi}:{col_p}{ff}")
            earnings_raw = self.sh_lp.get(f"{col_e}{fi}:{col_e}{ff}")

            resultado = {}
            for i, row_t in enumerate(tickers_raw):
                ticker = row_t[0].strip().upper() if row_t else ""
                if not ticker:
                    continue
                try:
                    p_raw = puntos_raw[i][0] if i < len(puntos_raw) and puntos_raw[i] else None
                    p_val = float(str(p_raw).replace('%','').strip()) if p_raw not in (None, '', 'N/A') else None
                    e_raw = earnings_raw[i][0].strip() if i < len(earnings_raw) and earnings_raw[i] else 'N/A'
                    resultado[ticker] = {'pts': p_val, 'earnings': e_raw}
                except:
                    resultado[ticker] = {'pts': None, 'earnings': 'N/A'}
            print(f"📋 Puntuaciones LP cargadas: {len(resultado)} empresas\n")
            return resultado
        except Exception as e:
            print(f"⚠️  Error leyendo LP: {e}")
            return {}

    # ──────────────────────────────────────────────────────────────
    # INDICADORES TÉCNICOS
    # ──────────────────────────────────────────────────────────────
    def _williams_r(self, df, p=14):
        hh = df['High'].rolling(p).max()
        ll = df['Low'].rolling(p).min()
        return ((hh - df['Close']) / (hh - ll + 1e-9)) * -100

    def _rsi(self, df, p=14):
        d = df['Close'].diff()
        g = d.clip(lower=0).rolling(p).mean()
        l = (-d.clip(upper=0)).rolling(p).mean()
        return 100 - (100 / (1 + g / l.replace(0, np.nan)))

    def _macd(self, df, f=12, s=26, sig=9):
        ef = df['Close'].ewm(span=f, adjust=False).mean()
        es = df['Close'].ewm(span=s, adjust=False).mean()
        ml = ef - es
        sl = ml.ewm(span=sig, adjust=False).mean()
        return ml, sl, ml - sl

    def _bollinger(self, df, p=20, k=2):
        sma   = df['Close'].rolling(p).mean()
        std   = df['Close'].rolling(p).std()
        upper = sma + k * std
        lower = sma - k * std
        pct_b = (df['Close'] - lower) / (upper - lower + 1e-9)
        width = (upper - lower) / (sma + 1e-9)
        return upper, sma, lower, pct_b, width

    def _atr(self, df, p=14):
        tr = pd.concat([
            df['High'] - df['Low'],
            (df['High'] - df['Close'].shift()).abs(),
            (df['Low']  - df['Close'].shift()).abs()
        ], axis=1).max(axis=1)
        return tr.rolling(p).mean()

    def _divergencia_wr(self, df, wr, lb=10):
        p = df['Close'].values; w = wr.values
        if len(p) < lb:
            return "Ninguna"
        pr, wr2 = p[-lb:], w[-lb:]
        if pr[-1] < np.min(pr[:-1]) and wr2[-1] > np.min(wr2[:-1]) and wr2[-1] < -50:
            return "🟢 Alcista (Bullish)"
        if pr[-1] > np.max(pr[:-1]) and wr2[-1] < np.max(wr2[:-1]) and wr2[-1] > -50:
            return "🔴 Bajista (Bearish)"
        return "Ninguna"

    # ──────────────────────────────────────────────────────────────
    # PERSISTENCIA WR EN SOBREVENTA
    # ──────────────────────────────────────────────────────────────
    def _semanas_en_sobreventa(self, wr_series, umbral=-80):
        """Cuenta cuántas velas CONSECUTIVAS el WR lleva <= umbral."""
        vals = wr_series.dropna().values[::-1]  # más reciente primero
        count = 0
        for v in vals:
            if v <= umbral:
                count += 1
            else:
                break
        return count

    # ──────────────────────────────────────────────────────────────
    # NIVELES FIBONACCI
    # ──────────────────────────────────────────────────────────────
    def _nivel_fibonacci(self, df, ventana=52):
        """
        Calcula retrocesos de Fibonacci sobre la ventana dada.
        Retorna (nivel_mas_cercano, distancia_pct, descripcion)
        """
        sub    = df['Close'].iloc[-ventana:]
        maximo = sub.max()
        minimo = sub.min()
        rango  = maximo - minimo
        precio = df['Close'].iloc[-1]

        if rango == 0:
            return None, 1.0, "Sin rango"

        niveles = {
            '23.6%': maximo - rango * 0.236,
            '38.2%': maximo - rango * 0.382,
            '50.0%': maximo - rango * 0.500,
            '61.8%': maximo - rango * 0.618,
            '78.6%': maximo - rango * 0.786,
        }
        # Nivel más cercano al precio actual
        mas_cercano = min(niveles.items(), key=lambda x: abs(x[1] - precio))
        dist_pct = abs(mas_cercano[1] - precio) / precio
        return mas_cercano[0], dist_pct, f"Fib {mas_cercano[0]} @ ${mas_cercano[1]:.2f}"

    # ──────────────────────────────────────────────────────────────
    # PATRONES DE VELA DE REVERSAL
    # ──────────────────────────────────────────────────────────────
    def _patron_reversal(self, df):
        """
        Detecta patrones de vela alcista de reversión en la última vela.
        Retorna (nombre_patron, fuerza: 0-2)
        """
        if len(df) < 2:
            return "Ninguno", 0

        curr = df.iloc[-1]; prev = df.iloc[-2]
        o, h, l, c = curr['Open'], curr['High'], curr['Low'], curr['Close']
        body   = abs(c - o)
        rango  = h - l
        if rango == 0:
            return "Ninguno", 0
        sombra_inf = min(o, c) - l
        sombra_sup = h - max(o, c)

        # Martillo / Hammer
        if (sombra_inf >= body * 2 and
                sombra_sup <= body * 0.3 and
                c > o and rango > 0):
            return "Martillo", 2

        # Martillo invertido alcista (tras caída)
        if (sombra_sup >= body * 2 and
                sombra_inf <= body * 0.3 and
                prev['Close'] < prev['Open']):
            return "Martillo Invertido", 1

        # Engulfing alcista
        if (c > o and                          # vela alcista
                prev['Close'] < prev['Open'] and   # vela anterior bajista
                c > prev['Open'] and               # cierre sobre apertura anterior
                o < prev['Close']):                # apertura bajo cierre anterior
            return "Engulfing Alcista", 2

        # Doji en zona baja (indecisión = posible giro)
        if body <= rango * 0.1 and rango > 0:
            return "Doji", 1

        # Vela alcista fuerte (cierre en 80%+ del rango)
        if c > o and (c - l) / rango >= 0.80:
            return "Vela Alcista Fuerte", 1

        return "Ninguno", 0

        # ──────────────────────────────────────────────────────────────
    # PUNTUAR UN TIMEFRAME
    # ──────────────────────────────────────────────────────────────
    def _puntuar_tf(self, df, label=""):
        pts = 0; señales = []

        # Williams %R
        w14 = self._williams_r(df, 14)
        w7  = self._williams_r(df, 7)
        wr  = w14.iloc[-1]; wr7 = w7.iloc[-1]
        if   wr <= -80: pts += 20; señales.append(f"[{label}] WR14 sobreventa extrema ({wr:.1f})")
        elif wr <= -60: pts += 10; señales.append(f"[{label}] WR14 sobreventa moderada ({wr:.1f})")
        elif wr >= -20: pts -= 15; señales.append(f"[{label}] WR14 sobrecompra ({wr:.1f})")
        if wr <= -70 and wr7 <= -70:
            pts += 10; señales.append(f"[{label}] WR multi-periodo alineado")

        # RSI
        rsi = self._rsi(df, 14); rv = rsi.iloc[-1]
        if   rv <= 30: pts += 15; señales.append(f"[{label}] RSI sobreventa ({rv:.1f})")
        elif rv <= 40: pts += 8;  señales.append(f"[{label}] RSI zona baja ({rv:.1f})")
        elif rv >= 70: pts -= 12; señales.append(f"[{label}] RSI sobrecompra ({rv:.1f})")
        elif rv >= 60: pts -= 5
        if len(df) >= 10:
            pa = df['Close'].values[-10:]; ra = rsi.values[-10:]
            if pa[-1] < np.min(pa[:-1]) and ra[-1] > np.min(ra[:-1]) and rv < 45:
                pts += 12; señales.append(f"[{label}] Divergencia alcista RSI")

        # MACD
        ml, ms, mh = self._macd(df)
        mv = ml.iloc[-1]; msv = ms.iloc[-1]; mhv = mh.iloc[-1]
        mp = mh.iloc[-2] if len(mh) > 1 else mhv
        if mv > msv and ml.iloc[-2] <= ms.iloc[-2]:
            pts += 15; señales.append(f"[{label}] Cruce alcista MACD")
        elif mhv > mp and mhv < 0:
            pts += 8;  señales.append(f"[{label}] MACD histograma mejorando")
        elif mv > msv:
            pts += 5;  señales.append(f"[{label}] MACD sobre señal")
        elif mv < msv and ml.iloc[-2] >= ms.iloc[-2]:
            pts -= 12; señales.append(f"[{label}] Cruce bajista MACD")

        # Bollinger
        bbu, bbm, bbl, pct_b, bbw = self._bollinger(df)
        pbv = pct_b.iloc[-1]; bwv = bbw.iloc[-1]
        bwm = bbw.rolling(10).mean().iloc[-1]
        if   pbv <= 0.05: pts += 15; señales.append(f"[{label}] Precio en/bajo banda inferior BB")
        elif pbv <= 0.20: pts += 8;  señales.append(f"[{label}] Precio cerca banda inferior BB")
        elif pbv >= 0.95: pts -= 12; señales.append(f"[{label}] Precio en/sobre banda superior BB")
        if bwv < bwm * 0.75:
            pts += 7; señales.append(f"[{label}] BB Squeeze detectado")

        # EMAs / SMAs
        precio = df['Close'].iloc[-1]
        ema8   = df['Close'].ewm(span=8,  adjust=False).mean().iloc[-1]
        ema21  = df['Close'].ewm(span=21, adjust=False).mean().iloc[-1]
        sma20  = df['Close'].rolling(20).mean().iloc[-1]
        sma50  = df['Close'].rolling(50).mean().iloc[-1] if len(df) >= 50 else ema21
        if precio > sma20: pts += 8;  señales.append(f"[{label}] Precio sobre SMA20")
        if ema8 > ema21:   pts += 5;  señales.append(f"[{label}] EMA8 > EMA21")
        if precio > sma50: pts += 5;  señales.append(f"[{label}] Precio sobre SMA50")

        # Divergencia WR
        div = self._divergencia_wr(df, w14)
        if "Alcista" in div: pts += 15; señales.append(f"[{label}] Divergencia alcista WR%")
        elif "Bajista" in div: pts -= 10; señales.append(f"[{label}] Divergencia bajista WR%")

        # Volumen
        vr = df['Volume'].iloc[-1] / df['Volume'].rolling(10).mean().iloc[-1]
        if   vr >= 1.5: pts += 8; señales.append(f"[{label}] Volumen elevado ({vr:.1f}x)")
        elif vr >= 1.2: pts += 4

        # ── CALIDAD DE SEÑAL: Persistencia en sobreventa ─────
        semanas_sv = self._semanas_en_sobreventa(w14, umbral=-80)
        if   semanas_sv >= 4:
            pts += 18; señales.append(f"[{label}] WR en sobreventa {semanas_sv} velas (señal madura)")
        elif semanas_sv >= 2:
            pts += 10; señales.append(f"[{label}] WR en sobreventa {semanas_sv} velas (consolidando)")
        elif semanas_sv == 1:
            pts += 4;  señales.append(f"[{label}] WR en sobreventa 1 vela (señal nueva)")

        # ── CALIDAD DE SEÑAL: Fibonacci ───────────────────────
        fib_nivel, fib_dist, fib_desc = self._nivel_fibonacci(df)
        if fib_nivel and fib_dist <= 0.02:    # dentro del 2% del nivel
            pts += 15; señales.append(f"[{label}] Precio en soporte Fibonacci {fib_desc}")
        elif fib_nivel and fib_dist <= 0.05:  # dentro del 5%
            pts += 8;  señales.append(f"[{label}] Precio cerca de Fibonacci {fib_desc}")

        # ── CALIDAD DE SEÑAL: Patrón de vela reversal ─────────
        patron, fuerza = self._patron_reversal(df)
        if   fuerza == 2:
            pts += 15; señales.append(f"[{label}] Patrón reversal fuerte: {patron}")
        elif fuerza == 1:
            pts += 7;  señales.append(f"[{label}] Patrón reversal moderado: {patron}")

        return {
            'puntos': pts, 'señales': señales,
            'wr14': wr, 'wr7': wr7, 'rsi': rv, 'macd_h': mhv, 'pct_b': pbv,
            'div': div, 'sma20': sma20, 'sma50': sma50, 'v_ratio': vr,
            'w14_s': w14, 'w7_s': w7, 'rsi_s': rsi,
            'macd_l': ml, 'macd_sig': ms, 'macd_h_s': mh,
            'bb_upper': bbu, 'bb_lower': bbl, 'bb_mid': bbm,
            # calidad
            'semanas_sv': semanas_sv,
            'fib_nivel': fib_nivel, 'fib_dist': fib_dist, 'fib_desc': fib_desc,
            'patron': patron, 'fuerza_patron': fuerza,
        }

    # ──────────────────────────────────────────────────────────────
    # ANÁLISIS COMPLETO DE UN TICKER
    # ──────────────────────────────────────────────────────────────
    def analizar(self, ticker_sym: str, pts_lp=None, earnings_window: str = 'N/A') -> dict | None:
        try:
            print(f"  📊 {ticker_sym}...", end=" ", flush=True)
            stock = yf.Ticker(ticker_sym)
            df_w  = stock.history(period="2y",  interval="1wk")
            df_d  = stock.history(period="6mo", interval="1d")

            if df_w.empty or len(df_w) < 30:
                print("❌ Datos insuficientes"); return None

            precio = df_w['Close'].iloc[-1]
            atr    = self._atr(df_w, 14).iloc[-1]

            # ── Puntuación semanal (60%) ──────────────────────────
            rw = self._puntuar_tf(df_w, "W")

            # ── Puntuación diaria (40%) ───────────────────────────
            if not df_d.empty and len(df_d) >= 30:
                rd = self._puntuar_tf(df_d, "D")
                confluencia = 15 if (rw['puntos'] > 0 and rd['puntos'] > 0) else \
                             -10 if (rw['puntos'] < 0 and rd['puntos'] < 0) else 0
                pts_swing_raw = rw['puntos'] * 0.60 + rd['puntos'] * 0.40 + confluencia
                señales_todas = rw['señales'] + rd['señales']
                wr_d = rd['wr14']; rsi_d = rd['rsi']; div_d = rd['div']
            else:
                rd = None
                pts_swing_raw = rw['puntos']
                señales_todas = rw['señales']
                wr_d = rsi_d = None; div_d = "N/D"

            # Normalizar swing a 0-100
            pts_swing = max(0, min(100, int((pts_swing_raw + 30) / 1.6)))

            # ── Matriz de decisión combinada ──────────────────────
            mx = matriz_decision(pts_swing, pts_lp)

            # Puntuación combinada para ranking (swing + bonus LP)
            pts_combo = max(0, min(100, pts_swing + mx['bonus']))

            # ── Earnings Risk (veto o penalización) ──────────────
            ew = earnings_window.strip() if earnings_window else 'N/A'
            VETO_EARNINGS = ew in ('ESTA SEMANA',)
            ALERTA_EARNINGS = ew in ('ESTE MES',)
            if VETO_EARNINGS:
                pts_combo = 0
                mx['accion'] = f'🚫 VETO EARNINGS ({ew}) — ' + mx['accion']
                mx['acc']    = 'ESPERAR'
                señales_todas.insert(0, f'⚠️ EARNINGS ESTA SEMANA — NO OPERAR')
            elif ALERTA_EARNINGS:
                pts_combo = max(0, pts_combo - 20)
                mx['accion'] = f'⚠️ EARNINGS PRÓX ({ew}) — ' + mx['accion']
                señales_todas.insert(0, f'⚠️ Earnings este mes — reducir tamaño de posición')

            # ── SL / TP dinámicos ─────────────────────────────────
            sl  = round(precio - atr * 2.0, 2)
            tp1 = round(precio + atr * 3.0, 2)
            tp2 = round(precio + atr * 5.0, 2)
            rb  = round((tp1 - precio) / max(precio - sl, 0.01), 2)

            setup = ("Setup Perfecto MTF+LP"
                     if rw['wr14'] <= -80 and "Alcista" in rw['div']
                        and mx['zona_lp'] == "ALTO"
                     else "Análisis Normal")

            sig = (f"EW:{ew} | WR(W):{rw['wr14']:.0f} RSI(W):{rw['rsi']:.0f}"
                   + (f" | WR(D):{wr_d:.0f} RSI(D):{rsi_d:.0f}" if wr_d else "")
                   + f" | LP:{pts_lp if pts_lp is not None else 'N/D'}")

            print(f"Swing:{pts_swing}  LP:{pts_lp}  EW:{ew}  SV:{rw['semanas_sv']}sem  Fib:{rw['fib_nivel']}  Pat:{rw['patron']}  Combo:{pts_combo}  → {mx['accion']}")

            return {
                'ticker':   ticker_sym,
                'precio':   round(precio, 2),
                # Semanal
                'r14':      round(rw['wr14'], 1),
                'r7':       round(rw['wr7'],  1),
                'rsi_w':    round(rw['rsi'],  1),
                'macd_w':   round(rw['macd_h'], 4),
                'pct_b_w':  round(rw['pct_b'] * 100, 1),
                'div':      rw['div'],
                'sma20':    round(rw['sma20'], 2),
                'sma50':    round(rw['sma50'], 2),
                'vol':      round(rw['v_ratio'], 2),
                'zona':     ("Sobreventa" if rw['wr14'] <= -80
                             else "Sobrecompra" if rw['wr14'] >= -20 else "Neutral"),
                # Diario
                'r14_d':    round(wr_d,  1) if wr_d  is not None else None,
                'rsi_d':    round(rsi_d, 1) if rsi_d is not None else None,
                'div_d':    div_d,
                # Puntuaciones
                'pts_swing': pts_swing,
                'pts_lp':    pts_lp if pts_lp is not None else 'N/D',
                'pts_combo': pts_combo,
                'zona_lp':   mx['zona_lp'],
                'zona_sw':   mx['zona_sw'],
                # Decisión combinada
                'dec':       mx['accion'],
                'acc':       mx['acc'],
                'rec':       setup,
                'ew':        ew,
                # calidad señal semanal
                'semanas_sv':   rw['semanas_sv'],
                'fib_nivel':    rw['fib_nivel'],
                'fib_dist':     round(rw['fib_dist'] * 100, 1),
                'fib_desc':     rw['fib_desc'],
                'patron':       rw['patron'],
                'fuerza_patron':rw['fuerza_patron'],
                'sig':       sig,
                # Risk management
                'atr':  round(atr, 2),
                'sl':   sl, 'tp1': tp1, 'tp2': tp2, 'rb': rb,
                # Series para gráfica
                'df_w': df_w, 'df_d': df_d,
                'rw':   rw,   'rd':   rd,
                'señales': señales_todas,
            }
        except Exception as e:
            print(f"❌ Error: {e}"); return None

    # ──────────────────────────────────────────────────────────────
    # GRÁFICA (5 paneles)
    # ──────────────────────────────────────────────────────────────
    def graficar(self, d: dict, guardar_como=None):
        ticker = d['ticker']
        df     = d['df_w'].copy().iloc[-52:]
        rw     = d['rw']

        w14      = rw['w14_s'].iloc[-52:]
        w7       = rw['w7_s'].iloc[-52:]
        rsi_s    = rw['rsi_s'].iloc[-52:]
        macd_l   = rw['macd_l'].iloc[-52:]
        macd_sig = rw['macd_sig'].iloc[-52:]
        macd_h   = rw['macd_h_s'].iloc[-52:]
        bb_up    = rw['bb_upper'].iloc[-52:]
        bb_lo    = rw['bb_lower'].iloc[-52:]
        bb_mi    = rw['bb_mid'].iloc[-52:]

        BG = '#0d1117'; PANEL = '#161b22'
        GR = '#26a641'; RD = '#f85149'; YL = '#e3b341'
        BL = '#58a6ff'; PU = '#bc8cff'; OR = '#f0883e'
        GY = '#8b949e'; WH = '#e6edf3'; TE = '#39d353'

        fig = plt.figure(figsize=(18, 16), facecolor=BG)
        gs  = gridspec.GridSpec(5, 1, figure=fig,
                                height_ratios=[3, 1.2, 1.2, 1.2, 1],
                                hspace=0.06)
        axes = [fig.add_subplot(gs[i]) for i in range(5)]
        for ax in axes:
            ax.set_facecolor(PANEL)
            ax.tick_params(colors=GY, labelsize=8)
            ax.spines[:].set_color('#30363d')

        fechas = df.index; n = len(df); xs = range(n)
        ax1, ax2, ax3, ax4, ax5 = axes

        # ── Panel 1: Velas + BB + SMAs ────────────────────────────
        for i, (_, row) in enumerate(df.iterrows()):
            c = GR if row['Close'] >= row['Open'] else RD
            ax1.plot([i,i],[row['Low'],row['High']], color=c, lw=0.8, alpha=0.7)
            ax1.add_patch(plt.Rectangle((i-0.3, min(row['Open'],row['Close'])),
                          0.6, abs(row['Close']-row['Open']), color=c, alpha=0.85))

        ax1.plot(xs, bb_up.values, color=PU, lw=0.9, ls='--', alpha=0.7, label='BB Sup')
        ax1.plot(xs, bb_mi.values, color=GY, lw=0.9, ls='-',  alpha=0.4, label='BB Med')
        ax1.plot(xs, bb_lo.values, color=TE, lw=0.9, ls='--', alpha=0.7, label='BB Inf')
        ax1.fill_between(xs, bb_lo.values, bb_up.values, alpha=0.04, color=BL)
        sma20v = df['Close'].rolling(20).mean()
        sma50v = df['Close'].rolling(50).mean()
        ax1.plot(xs, sma20v.values, color=BL,   lw=1.3, label='SMA20')
        ax1.plot(xs, sma50v.values, color=PU,   lw=1.3, label='SMA50', ls='--')
        ax1.axhline(d['precio'], color=YL, lw=1.5, label=f"Entrada ${d['precio']}")
        ax1.axhline(d['sl'],     color=RD, lw=1.5, ls='--', label=f"SL ${d['sl']}")
        ax1.axhline(d['tp1'],    color=GR, lw=1.5, ls='--', label=f"TP1 ${d['tp1']}")
        ax1.axhline(d['tp2'],    color=GR, lw=2.0, label=f"TP2 ${d['tp2']}")
        ax1.axhspan(d['sl'], d['precio'], alpha=0.07, color=RD)
        ax1.axhspan(d['precio'], d['tp1'], alpha=0.05, color=GR)
        ax1.axhspan(d['tp1'], d['tp2'],   alpha=0.09, color=GR)
        for val, lbl, col in [(d['tp2'], f"TP2 ${d['tp2']}", GR),
                               (d['tp1'], f"TP1 ${d['tp1']}", GR),
                               (d['precio'], f"ENTRADA ${d['precio']}", YL),
                               (d['sl'], f"SL ${d['sl']}", RD)]:
            ax1.annotate(lbl, xy=(n-1, val), xytext=(n+1, val), color=col,
                         fontsize=8, fontweight='bold', va='center',
                         arrowprops=dict(arrowstyle='->', color=col, lw=1))
        ax1.set_xlim(-1, n+10)
        rb_c = GR if d['rb'] >= 2 else YL if d['rb'] >= 1.5 else RD
        ax1.text(0.01, 0.97,
                 f"R:R={d['rb']}x  Riesgo:${round(d['precio']-d['sl'],2)}"
                 f"  TP1:+${round(d['tp1']-d['precio'],2)}",
                 transform=ax1.transAxes, color=rb_c, fontsize=8.5, va='top',
                 fontweight='bold',
                 bbox=dict(boxstyle='round,pad=0.3', facecolor=BG, alpha=0.7))

        # Etiqueta de puntuaciones LP + Swing en la gráfica
        lp_str = f"{d['pts_lp']}" if d['pts_lp'] != 'N/D' else "N/D"
        lp_c   = GR if d['zona_lp'] == 'ALTO' else YL if d['zona_lp'] == 'MEDIO' else RD
        ax1.text(0.99, 0.97,
                 f"Swing: {d['pts_swing']}/100  |  LP: {lp_str}  |  Zona LP: {d['zona_lp']}",
                 transform=ax1.transAxes, color=lp_c, fontsize=8.5, va='top', ha='right',
                 fontweight='bold',
                 bbox=dict(boxstyle='round,pad=0.3', facecolor=BG, alpha=0.7))

        ax1.legend(loc='upper left', fontsize=7, facecolor=PANEL,
                   labelcolor=WH, framealpha=0.8, ncol=5)
        ax1.set_ylabel('Precio (USD)', color=GY, fontsize=9)
        ax1.tick_params(labelbottom=False)

        # ── Panel 2: Williams %R ──────────────────────────────────
        ax2.plot(xs, w14.values, color=BL, lw=1.5, label='W%R 14')
        ax2.plot(xs, w7.values,  color=OR, lw=1.0, label='W%R 7',  ls='--', alpha=0.75)
        ax2.axhline(-20, color=RD, lw=0.8, ls=':', alpha=0.7)
        ax2.axhline(-50, color=GY, lw=0.8, ls=':', alpha=0.5)
        ax2.axhline(-80, color=GR, lw=0.8, ls=':', alpha=0.7)
        ax2.axhspan(-100, -80, alpha=0.10, color=GR)
        ax2.axhspan(-20, 0,    alpha=0.10, color=RD)
        ax2.set_ylim(-105, 5)
        wv = w14.iloc[-1]; wc = GR if wv <= -80 else RD if wv >= -20 else YL
        ax2.scatter([n-1], [wv], color=wc, s=60, zorder=5)
        ax2.text(n-1, wv+3, f'{wv:.1f}', color=wc, fontsize=7.5, ha='center', fontweight='bold')
        ax2.set_ylabel('Williams %R', color=GY, fontsize=9)
        ax2.legend(loc='upper right', fontsize=7.5, facecolor=PANEL, labelcolor=WH, framealpha=0.8)
        ax2.tick_params(labelbottom=False)

        # ── Panel 3: RSI ──────────────────────────────────────────
        ax3.plot(xs, rsi_s.values, color=TE, lw=1.5, label='RSI 14')
        ax3.axhline(70, color=RD, lw=0.8, ls=':', alpha=0.7)
        ax3.axhline(50, color=GY, lw=0.8, ls=':', alpha=0.4)
        ax3.axhline(30, color=GR, lw=0.8, ls=':', alpha=0.7)
        ax3.axhspan(0, 30,   alpha=0.08, color=GR)
        ax3.axhspan(70, 100, alpha=0.08, color=RD)
        ax3.set_ylim(0, 100)
        rv2 = rsi_s.iloc[-1]; rc = GR if rv2 <= 30 else RD if rv2 >= 70 else YL
        ax3.scatter([n-1], [rv2], color=rc, s=60, zorder=5)
        ax3.text(n-1, rv2+2, f'{rv2:.1f}', color=rc, fontsize=7.5, ha='center', fontweight='bold')
        ax3.set_ylabel('RSI', color=GY, fontsize=9)
        ax3.legend(loc='upper right', fontsize=7.5, facecolor=PANEL, labelcolor=WH, framealpha=0.8)
        ax3.tick_params(labelbottom=False)

        # ── Panel 4: MACD ─────────────────────────────────────────
        ax4.plot(xs, macd_l.values,   color=BL,   lw=1.3, label='MACD')
        ax4.plot(xs, macd_sig.values, color=OR,   lw=1.0, label='Señal', ls='--')
        for i, h in enumerate(macd_h.values):
            ax4.bar(i, h, color=GR if h >= 0 else RD, alpha=0.6, width=0.8)
        ax4.axhline(0, color=GY, lw=0.8, alpha=0.4)
        ax4.set_ylabel('MACD', color=GY, fontsize=9)
        ax4.legend(loc='upper right', fontsize=7.5, facecolor=PANEL, labelcolor=WH, framealpha=0.8)
        ax4.tick_params(labelbottom=False)

        # ── Panel 5: Volumen ──────────────────────────────────────
        vm = df['Volume'].rolling(10).mean()
        for i, (v, vm_) in enumerate(zip(df['Volume'].values, vm.values)):
            ax5.bar(i, v, color=GR if v > vm_ else GY, alpha=0.75, width=0.8)
        ax5.plot(xs, vm.values, color=YL, lw=1.0, ls='--', label='Vol MA10')
        ax5.set_ylabel('Volumen', color=GY, fontsize=9)
        ax5.legend(loc='upper right', fontsize=7.5, facecolor=PANEL, labelcolor=WH, framealpha=0.8)

        step = max(1, n // 8)
        ticks = list(range(0, n, step))
        ax5.set_xticks(ticks)
        ax5.set_xticklabels([fechas[i].strftime('%b %Y') for i in ticks],
                            color=GY, fontsize=8, rotation=30)

        # ── Título ────────────────────────────────────────────────
        cc = GR if 'COMPRA' in d['dec'] else YL if 'ESPERAR' in d['dec'] or 'SCALP' in d['dec'] else RD
        rsi_d_str = f"RSI(D):{d['rsi_d']}" if d['rsi_d'] else ""
        fig.suptitle(
            f"{ticker}  —  {d['dec']}  |  Combo: {d['pts_combo']}/100\n"
            f"Swing: {d['pts_swing']}/100  •  LP: {d['pts_lp']}  •  Zona LP: {d['zona_lp']}  •  "
            f"WR(W):{d['r14']} RSI(W):{d['rsi_w']}  |  "
            f"WR(D):{d.get('r14_d','N/D')} {rsi_d_str}  |  Div: {d['div']}",
            color=WH, fontsize=11, fontweight='bold', y=0.99,
            bbox=dict(boxstyle='round,pad=0.4', facecolor=cc, alpha=0.25)
        )
        plt.tight_layout(rect=[0, 0, 1, 0.97])

        nombre = guardar_como or f"swing_{ticker}_{datetime.now().strftime('%Y%m%d_%H%M')}.png"
        plt.savefig(nombre, dpi=150, bbox_inches='tight', facecolor=BG)
        plt.close()
        print(f"      📈 Gráfica: {nombre}")
        return nombre

    # ──────────────────────────────────────────────────────────────
    # ACTUALIZAR HOJA + GENERAR TOP-N
    # ──────────────────────────────────────────────────────────────
    def ejecutar(self):
        cfg = self.cfg

        # 1. Leer puntuaciones largo plazo
        scores_lp = self.leer_puntuaciones_lp()

        # 2. Leer tickers del swing
        col_data = self.sh_swing.col_values(cfg['swing_col_tickers'])
        ff = cfg['swing_fila_fin'] or len(col_data)
        fi = cfg['swing_fila_ini']
        tickers = [t.strip().upper() for t in col_data[fi-1:ff] if t.strip()]

        if not tickers:
            return print("⚠️  No hay tickers.")

        print(f"🔍 {len(tickers)} tickers a analizar\n{'='*60}")

        resultados = []; filas = []

        for t in tickers:
            lp_data   = scores_lp.get(t, {'pts': None, 'earnings': 'N/A'})
            pts_lp    = lp_data['pts']
            ew        = lp_data.get('earnings', 'N/A')   # Earnings Window
            d = self.analizar(t, pts_lp, ew)
            if d:
                fila = [
                    safe(d['r14']),    safe(d['r7']),    safe(d['rsi_w']),
                    safe(d['macd_w']), safe(d['pct_b_w']),
                    d['zona'], d['div'],
                    safe(d['r14_d']) if d['r14_d'] else 'N/D',
                    safe(d['rsi_d']) if d['rsi_d'] else 'N/D',
                    d['div_d'],
                    safe(d['sma20']), safe(d['sma50']), safe(d['vol']), safe(d['atr']),
                    safe(d['pts_swing']),   # puntuación swing sola
                    safe(d['pts_lp']) if d['pts_lp'] != 'N/D' else 'N/D',  # LP
                    safe(d['pts_combo']),   # puntuación combinada ← ranking
                    d['zona_lp'], d['zona_sw'],
                    d['dec'], d['acc'], d['rec'], d['sig'], d['ew'],
                    safe(d['semanas_sv']), d['fib_nivel'] or 'N/D',
                    safe(d['fib_dist']), d['patron'],
                    safe(d['sl']), safe(d['tp1']), safe(d['tp2']), safe(d['rb'])
                ]
                filas.append(fila)
                resultados.append(d)
            else:
                filas.append(['ERROR'] * 32)

        # 3. Escribir en Google Sheets (27 columnas)
        if filas:
            n_cols = len(filas[0]) if filas else 27
            ci = letter_to_col(cfg['swing_col_salida'])
            ce = col_to_letter(ci + n_cols - 1)
            rango = f"{cfg['swing_col_salida']}{fi}:{ce}{fi + len(filas) - 1}"
            self.sh_swing.update(range_name=rango, values=filas)
            print(f"\n✅ Datos escritos en {rango}")

        # 4. Top N por pts_combo → gráficas
        if resultados:
            top = sorted(resultados, key=lambda x: x['pts_combo'], reverse=True)[:cfg['top_n']]
            print(f"\n{'='*60}")
            print(f"  🏆  TOP {cfg['top_n']} MEJORES SETUPS COMBINADOS")
            print(f"{'='*60}")
            for rank, d in enumerate(top, 1):
                print(f"\n  #{rank}  {d['ticker']}  "
                      f"Combo:{d['pts_combo']} Swing:{d['pts_swing']} LP:{d['pts_lp']}  "
                      f"→  {d['dec']}")
                self.graficar(d, f"TOP{rank}_{d['ticker']}_{datetime.now().strftime('%Y%m%d_%H%M')}.png")

            # 5. Resumen de señales top 3
            print(f"\n{'─'*60}")
            print("  📋  SEÑALES DETALLADAS — TOP 3")
            print(f"{'─'*60}")
            for d in top[:3]:
                print(f"\n  {d['ticker']}  (Combo:{d['pts_combo']} | Swing:{d['pts_swing']} | LP:{d['pts_lp']}):")
                for s in d['señales']:
                    print(f"    • {s}")


# ══════════════════════════════════════════════════════════════════
# PUNTO DE ENTRADA
# ══════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    sistema = SwingSystemV3(CFG)
    sistema.ejecutar()
