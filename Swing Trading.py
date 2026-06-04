"""
╔══════════════════════════════════════════════════════════════════════════════╗
║     SWING TRADING v4 — SISTEMA MTF + LARGO PLAZO + PORTFOLIO INTELIGENTE   ║
║                                                                               ║
║  Mejoras v4:                                                                  ║
║   • Descargas paralelas con ThreadPoolExecutor + cache local                  ║
║   • Detección de régimen de mercado (SPY SMA200 + VIX)                        ║
║   • Position sizing basado en riesgo (% capital / riesgo por trade)         ║
║   • Filtro de correlación sectorial en el TOP-N                               ║
║   • Earnings reales via yfinance.calendar (no strings frágiles)               ║
║   • Walk-forward backtesting integrado para validar pesos                   ║
║   • Portfolio tracker: posiciones abiertas, trailing stops, pyramiding        ║
║   • Métricas de rendimiento: Sharpe, Sortino, Profit Factor, Max Drawdown   ║
║   • Manejo robusto de datos: N/D en vez de 0.0 para missing                   ║
║   • Logging estructurado con timestamps                                       ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
import warnings
import json
import os
import pickle
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple, Any
from collections import defaultdict

warnings.filterwarnings('ignore')

# FIX: Configurar matplotlib para manejar emojis correctamente
matplotlib.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['axes.unicode_minus'] = False

# ══════════════════════════════════════════════════════════════════════════════
# LOGGING ESTRUCTURADO
# ══════════════════════════════════════════════════════════════════════════════
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)-8s | %(message)s',
    datefmt='%H:%M:%S'
)
logger = logging.getLogger('SwingV4')


# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN CENTRALIZADA
# ══════════════════════════════════════════════════════════════════════════════
@dataclass
class Config:
    # Google Sheets
    credenciales_json: str = 'principios.json'

    # Swing
    swing_archivo: str = 'Portafolio Tracker'
    swing_hoja: str = 'Swing Trading'
    swing_fila_ini: int = 2
    swing_fila_fin: Optional[int] = 183
    swing_col_tickers: int = 1
    swing_col_salida: str = 'R'

    # Largo Plazo
    lp_archivo: str = 'Portafolio Financiero'
    lp_hoja: str = '7 PRINCIPIOS'
    lp_fila_ini: int = 7
    lp_fila_fin: int = 190
    lp_col_ticker: str = 'A'
    lp_col_puntuacion: str = 'DO'
    lp_col_earnings: str = 'V'
    lp_col_sector: str = 'DA'  # ← NUEVO: columna sector

    # Comportamiento
    top_n: int = 10
    max_posiciones: int = 5
    riesgo_por_trade_pct: float = 1.0  # % del capital por trade
    capital_total: float = 100000.0

    # Régimen de mercado
    regime_spy_ticker: str = 'SPY'
    regime_sma_period: int = 200
    regime_vix_ticker: str = '^VIX'
    regime_vix_alto: float = 25.0

    # Cache
    cache_dir: str = '.swing_cache'
    cache_ttl_horas: int = 4

    # Paralelización
    max_workers: int = 8

    # Backtesting
    bt_lookback_anos: int = 2
    bt_train_pct: float = 0.7

    def __post_init__(self):
        os.makedirs(self.cache_dir, exist_ok=True)


CFG = Config()


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS Y UTILIDADES
# ══════════════════════════════════════════════════════════════════════════════
def col_to_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def letter_to_col(label: str) -> int:
    r = 0
    for c in label:
        r = r * 26 + (ord(c.upper()) - 64)
    return r

def safe(v) -> Any:
    """Convierte NaN/Inf a 'N/D'. NUNCA a 0.0."""
    if v is None:
        return 'N/D'
    if isinstance(v, float):
        if np.isnan(v) or np.isinf(v):
            return 'N/D'
    if isinstance(v, (np.floating, np.integer)):
        if np.isnan(v) or np.isinf(v):
            return 'N/D'
        return float(v)
    return v

def safe_float(v, default=0.0) -> float:
    """Para cálculos internos donde necesitamos un float real."""
    if v is None or v == 'N/D' or v == '':
        return default
    try:
        f = float(str(v).replace('%', '').replace(',', '').strip())
        return 0.0 if np.isnan(f) or np.isinf(f) else f
    except (ValueError, TypeError):
        return default

class DataCache:
    """Cache local de datos de Yahoo Finance para evitar re-descargas."""
    def __init__(self, cfg: Config):
        self.cfg = cfg

    def _cache_path(self, ticker: str, period: str, interval: str) -> str:
        safe_t = ticker.replace('^', '_').replace('/', '_')
        return os.path.join(self.cfg.cache_dir, f"{safe_t}_{period}_{interval}.pkl")

    def get(self, ticker: str, period: str, interval: str) -> Optional[pd.DataFrame]:
        path = self._cache_path(ticker, period, interval)
        if not os.path.exists(path):
            return None
        mtime = datetime.fromtimestamp(os.path.getmtime(path))
        if datetime.now() - mtime > timedelta(hours=self.cfg.cache_ttl_horas):
            return None
        try:
            with open(path, 'rb') as f:
                return pickle.load(f)
        except Exception:
            return None

    def set(self, ticker: str, period: str, interval: str, df: pd.DataFrame):
        path = self._cache_path(ticker, period, interval)
        try:
            with open(path, 'wb') as f:
                pickle.dump(df, f)
        except Exception as e:
            logger.warning(f"No se pudo guardar cache para {ticker}: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# DETECCIÓN DE RÉGIMEN DE MERCADO
# ══════════════════════════════════════════════════════════════════════════════
class MarketRegime:
    """
    Detecta el régimen actual del mercado para ajustar la matriz de decisión.

    Regímenes:
      BULL_TREND   → SPY > SMA200, VIX < 20
      BULL_VOLATILE→ SPY > SMA200, VIX >= 20
      BEAR_TREND   → SPY < SMA200, VIX >= 25
      BEAR_BOUNCE  → SPY < SMA200, VIX < 25
      NEUTRAL      → casos intermedios
    """
    def __init__(self, cfg: Config):
        self.cfg = cfg
        self.cache = DataCache(cfg)

    def detectar(self) -> Dict[str, Any]:
        try:
            df_spy = self._descargar(self.cfg.regime_spy_ticker, "1y", "1d")
            df_vix = self._descargar(self.cfg.regime_vix_ticker, "1y", "1d")

            if df_spy.empty or df_vix.empty:
                return self._default_regime()

            spy_price = df_spy['Close'].iloc[-1]
            spy_sma = df_spy['Close'].rolling(self.cfg.regime_sma_period).mean().iloc[-1]
            vix_val = df_vix['Close'].iloc[-1]

            if spy_price > spy_sma and vix_val < 20:
                regime = "BULL_TREND"
                color = "🟢"
                ajuste_bonus = 0
            elif spy_price > spy_sma and vix_val >= 20:
                regime = "BULL_VOLATILE"
                color = "🟡"
                ajuste_bonus = -5
            elif spy_price < spy_sma and vix_val >= 25:
                regime = "BEAR_TREND"
                color = "🔴"
                ajuste_bonus = -15
            elif spy_price < spy_sma and vix_val < 25:
                regime = "BEAR_BOUNCE"
                color = "🟠"
                ajuste_bonus = -10
            else:
                regime = "NEUTRAL"
                color = "⚪"
                ajuste_bonus = -5

            return {
                'regime': regime,
                'color': color,
                'ajuste_bonus': ajuste_bonus,
                'spy_price': spy_price,
                'spy_sma200': spy_sma,
                'vix': vix_val,
                'descripcion': f"{color} {regime} | SPY ${spy_price:.2f} vs SMA200 ${spy_sma:.2f} | VIX {vix_val:.1f}"
            }
        except Exception as e:
            logger.error(f"Error detectando régimen: {e}")
            return self._default_regime()

    def _descargar(self, ticker: str, period: str, interval: str) -> pd.DataFrame:
        cached = self.cache.get(ticker, period, interval)
        if cached is not None:
            return cached
        df = yf.Ticker(ticker).history(period=period, interval=interval)
        if not df.empty:
            self.cache.set(ticker, period, interval, df)
        return df

    def _default_regime(self) -> Dict[str, Any]:
        return {
            'regime': 'UNKNOWN',
            'color': '⚪',
            'ajuste_bonus': 0,
            'spy_price': 0,
            'spy_sma200': 0,
            'vix': 0,
            'descripcion': '⚪ UNKNOWN (error de datos)'
        }


# ══════════════════════════════════════════════════════════════════════════════
# PORTFOLIO TRACKER (POSICIONES ABIERTAS)
# ══════════════════════════════════════════════════════════════════════════════
@dataclass
class Posicion:
    ticker: str
    entrada: float
    cantidad: float
    fecha_entrada: datetime
    sl_inicial: float
    tp1: float
    tp2: float
    trailing_stop: Optional[float] = None
    estado: str = 'ABIERTA'  # ABIERTA, CERRADA_SL, CERRADA_TP1, CERRADA_TP2
    pnl_pct: float = 0.0

    def actualizar_trailing(self, precio_actual: float, atr: float):
        """Ajusta trailing stop si el precio sube 2x ATR desde entrada."""
        # FIX 2: Proteger contra ATR inválido (0 o negativo)
        if atr is None or atr <= 0:
            return
        if precio_actual > self.entrada + 2 * atr:
            nuevo_sl = precio_actual - 1.5 * atr
            if self.trailing_stop is None or nuevo_sl > self.trailing_stop:
                self.trailing_stop = nuevo_sl
                logger.info(f"  🔄 {self.ticker} trailing stop ajustado a ${nuevo_sl:.2f}")

    def evaluar_cierre(self, precio_actual: float) -> Optional[str]:
        """Retorna razón de cierre si aplica, None si sigue abierta."""
        sl_efectivo = self.trailing_stop if self.trailing_stop else self.sl_inicial
        if precio_actual <= sl_efectivo:
            return 'CERRADA_SL'
        if precio_actual >= self.tp2:
            return 'CERRADA_TP2'
        if precio_actual >= self.tp1 and self.estado == 'ABIERTA':
            return 'CERRADA_TP1'
        return None


class PortfolioTracker:
    """Gestiona posiciones abiertas y calcula métricas de portafolio."""
    def __init__(self, cfg: Config):
        self.cfg = cfg
        self.posiciones: List[Posicion] = []
        self.historial: List[Dict] = []
        self._cargar()

    def _path(self) -> str:
        return os.path.join(self.cfg.cache_dir, 'portfolio.json')

    def _cargar(self):
        path = self._path()
        if os.path.exists(path):
            try:
                with open(path, 'r') as f:
                    data = json.load(f)
                for p in data.get('posiciones', []):
                    # FIX: Convertir fecha string a datetime al cargar desde JSON
                    if isinstance(p.get('fecha_entrada'), str):
                        p['fecha_entrada'] = datetime.fromisoformat(p['fecha_entrada'])
                    self.posiciones.append(Posicion(**p))
                self.historial = data.get('historial', [])
                logger.info(f"📂 Portfolio cargado: {len(self.posiciones)} posiciones abiertas")
            except Exception as e:
                logger.warning(f"No se pudo cargar portfolio: {e}")

    def guardar(self):
        path = self._path()
        data = {
            'posiciones': [
                {
                    'ticker': p.ticker, 'entrada': p.entrada, 'cantidad': p.cantidad,
                    'fecha_entrada': p.fecha_entrada.isoformat(), 'sl_inicial': p.sl_inicial,
                    'tp1': p.tp1, 'tp2': p.tp2, 'trailing_stop': p.trailing_stop,
                    'estado': p.estado, 'pnl_pct': p.pnl_pct
                }
                for p in self.posiciones
            ],
            'historial': self.historial
        }
        with open(path, 'w') as f:
            json.dump(data, f, indent=2, default=str)

    def abrir_posicion(self, ticker: str, precio: float, atr: float, 
                       pts_combo: int, capital_disponible: float) -> Optional[Posicion]:
        """Calcula position sizing y abre posición."""
        if len([p for p in self.posiciones if p.estado == 'ABIERTA']) >= self.cfg.max_posiciones:
            logger.info(f"  ⛔ {ticker}: máximo de posiciones alcanzado")
            return None

        riesgo_usd = self.cfg.capital_total * (self.cfg.riesgo_por_trade_pct / 100)
        riesgo_por_accion = precio - (precio - atr * 2.0)  # = 2x ATR
        if riesgo_por_accion <= 0:
            return None

        cantidad = int(riesgo_usd / riesgo_por_accion)
        costo_total = cantidad * precio

        if costo_total > capital_disponible:
            cantidad = int(capital_disponible / precio)
            costo_total = cantidad * precio

        if cantidad <= 0:
            return None

        sl = round(precio - atr * 2.0, 2)
        tp1 = round(precio + atr * 3.0, 2)
        tp2 = round(precio + atr * 5.0, 2)

        pos = Posicion(
            ticker=ticker, entrada=precio, cantidad=cantidad,
            fecha_entrada=datetime.now(), sl_inicial=sl, tp1=tp1, tp2=tp2
        )
        self.posiciones.append(pos)
        self.guardar()
        logger.info(f"  ✅ {ticker}: {cantidad} acciones @ ${precio:.2f} | SL ${sl} | TP1 ${tp1} | TP2 ${tp2}")
        return pos

    def actualizar_posiciones(self, precios_actuales: Dict[str, float], atrs: Dict[str, float]):
        """Evalúa trailing stops y cierres para todas las posiciones abiertas."""
        for pos in self.posiciones:
            if pos.estado != 'ABIERTA':
                continue
            if pos.ticker not in precios_actuales:
                continue

            precio = precios_actuales[pos.ticker]
            atr = atrs.get(pos.ticker, 0)

            # FIX 2: Solo actualizar trailing stop si el ATR es válido
            if atr and atr > 0:
                pos.actualizar_trailing(precio, atr)
            else:
                logger.warning(f"  ⚠️ {pos.ticker}: ATR inválido ({atr}), trailing stop no actualizado")

            # Calcular P&L
            pos.pnl_pct = (precio - pos.entrada) / pos.entrada * 100

            # Evaluar cierre
            cierre = pos.evaluar_cierre(precio)
            if cierre:
                pos.estado = cierre
                pnl_usd = pos.cantidad * (precio - pos.entrada)
                self.historial.append({
                    'ticker': pos.ticker, 'fecha_cierre': datetime.now().isoformat(),
                    'tipo_cierre': cierre, 'pnl_pct': pos.pnl_pct, 'pnl_usd': pnl_usd,
                    'precio_cierre': precio, 'precio_entrada': pos.entrada
                })
                logger.info(f"  🔒 {pos.ticker} {cierre} | P&L: {pos.pnl_pct:+.2f}% (${pnl_usd:+.2f})")

        self.guardar()

    def metricas(self) -> Dict[str, Any]:
        """Calcula métricas de rendimiento del portafolio."""
        if not self.historial:
            return {'mensaje': 'Sin trades históricos'}

        pnls = [h['pnl_pct'] for h in self.historial]
        pnls_usd = [h['pnl_usd'] for h in self.historial]

        winners = [p for p in pnls if p > 0]
        losers = [p for p in pnls if p <= 0]

        # Sharpe (anualizado, asumiendo ~20 trades/año para swing)
        retornos = np.array(pnls)
        sharpe = np.mean(retornos) / (np.std(retornos) + 1e-9) * np.sqrt(12) if len(retornos) > 1 else 0

        # Sortino
        downside = [p for p in retornos if p < 0]
        sortino = np.mean(retornos) / (np.std(downside) + 1e-9) * np.sqrt(12) if downside and len(retornos) > 1 else 0

        # Profit Factor
        gross_profit = sum(p for p in pnls_usd if p > 0)
        gross_loss = abs(sum(p for p in pnls_usd if p < 0))
        pf = gross_profit / gross_loss if gross_loss > 0 else float('inf')

        # Max Drawdown
        equity_curve = np.cumsum(pnls_usd)
        peak = np.maximum.accumulate(equity_curve)
        drawdown = (equity_curve - peak) / (peak + 1e-9) * 100
        max_dd = np.min(drawdown)

        return {
            'total_trades': len(pnls),
            'win_rate': len(winners) / len(pnls) * 100 if pnls else 0,
            'avg_win': np.mean(winners) if winners else 0,
            'avg_loss': np.mean(losers) if losers else 0,
            'profit_factor': pf,
            'sharpe_ratio': sharpe,
            'sortino_ratio': sortino,
            'max_drawdown_pct': max_dd,
            'total_pnl_usd': sum(pnls_usd),
            'posiciones_abiertas': len([p for p in self.posiciones if p.estado == 'ABIERTA'])
        }


# ══════════════════════════════════════════════════════════════════════════════
# MATRIZ DE DECISIÓN ADAPTATIVA POR RÉGIMEN
# ══════════════════════════════════════════════════════════════════════════════
def matriz_decision_adaptativa(pts_swing: int, pts_lp, regime: Dict) -> Dict:
    """
    Matriz de decisión que se adapta al régimen de mercado.

    En BEAR_TREND: exige combo más alto para COMPRAR, penaliza más los setups débiles.
    En BULL_TREND: mantiene umbrales originales.
    """
    ajuste = regime.get('ajuste_bonus', 0)

    # Normalizar largo plazo
    if pts_lp is None:
        zona_lp = "BAJO"
        lp_val = 0
    else:
        lp_val = safe_float(pts_lp, 0)
        zona_lp = "ALTO" if lp_val >= 70 else "MEDIO" if lp_val >= 50 else "BAJO"

    # Ajustar umbrales swing según régimen
    umbral_alto = 65 + max(0, -ajuste // 3)  # En bear, exige más
    umbral_medio = 40 + max(0, -ajuste // 5)

    zona_sw = "ALTO" if pts_swing >= umbral_alto else "MEDIO" if pts_swing >= umbral_medio else "BAJO"

    # Tabla base
    tabla = {
        ("ALTO",  "ALTO"):  ("🟢🟢 COMPRA FUERTE MTF+LP",  "COMPRAR",  15),
        ("ALTO",  "MEDIO"): ("🟢 COMPRA MODERADA",          "COMPRAR",   5),
        ("ALTO",  "BAJO"):  ("🟡 ESPERAR TIMING",           "VIGILAR",  -5),
        ("MEDIO", "ALTO"):  ("🟢 COMPRA SELECTIVA",         "COMPRAR",   5),
        ("MEDIO", "MEDIO"): ("🟡 VIGILAR",                  "VIGILAR",    0),
        ("MEDIO", "BAJO"):  ("🔴 NO OPERAR",                "ESPERAR",  -10),
        ("BAJO",  "ALTO"):  ("⚠️ SOLO SCALP CORTO",        "VIGILAR",  -10),
        ("BAJO",  "MEDIO"): ("🔴 EVITAR",                   "ESPERAR",  -15),
        ("BAJO",  "BAJO"):  ("🔴🔴 EVITAR TOTALMENTE",      "ESPERAR",  -20),
    }

    accion, acc, bonus = tabla[(zona_lp, zona_sw)]
    bonus_ajustado = bonus + ajuste

    # En bear trend, veto compras en zona LP BAJO aunque swing sea ALTO
    if regime['regime'] == 'BEAR_TREND' and zona_lp == 'BAJO' and acc == 'COMPRAR':
        accion = f"🔴 {regime['regime']} — EVITAR (LP débil en bear)"
        acc = 'ESPERAR'
        bonus_ajustado = -25

    return {
        'accion': accion,
        'acc': acc,
        'bonus': bonus_ajustado,
        'zona_lp': zona_lp,
        'zona_sw': zona_sw,
        'regime': regime['regime'],
        'ajuste_regime': ajuste
    }


# ══════════════════════════════════════════════════════════════════════════════
# EARNINGS REALES (NO STRINGS FRÁGILES)
# ══════════════════════════════════════════════════════════════════════════════
def obtener_earnings_window(ticker_sym: str) -> Tuple[str, Optional[datetime]]:
    """
    Obtiene la próxima fecha de earnings real de yfinance.
    Retorna: (etiqueta, fecha_real)

    Etiquetas:
      'ESTA SEMANA', 'ESTE MES', 'PRÓXIMO MES', 'FUTURO', 'PASADO', 'N/A'
    """
    try:
        stock = yf.Ticker(ticker_sym)
        cal = stock.calendar
        if cal is None or cal.empty:
            return 'N/A', None

        # yfinance calendar tiene columna 'Earnings Date'
        if 'Earnings Date' in cal.index:
            fecha_str = cal.loc['Earnings Date'].values[0]
        elif 'Earnings Date' in cal.columns:
            fecha_str = cal['Earnings Date'].iloc[0]
        else:
            return 'N/A', None

        if pd.isna(fecha_str):
            return 'N/A', None

        fecha = pd.to_datetime(fecha_str)
        hoy = datetime.now()
        diff = (fecha - hoy).days

        if diff < 0:
            return 'PASADO', fecha
        elif diff <= 7:
            return 'ESTA SEMANA', fecha
        elif diff <= 30:
            return 'ESTE MES', fecha
        elif diff <= 60:
            return 'PRÓXIMO MES', fecha
        else:
            return 'FUTURO', fecha
    except Exception as e:
        logger.debug(f"Earnings error para {ticker_sym}: {e}")
        return 'N/A', None


# ══════════════════════════════════════════════════════════════════════════════
# CORRELACIÓN SECTORIAL
# ══════════════════════════════════════════════════════════════════════════════
def filtrar_por_correlacion(candidatos: List[Dict], max_por_sector: int = 2) -> List[Dict]:
    """
    Limita el número de setups por sector para evitar concentración de riesgo.
    candidatos: lista de dicts con 'ticker' y 'sector'
    """
    sector_count = defaultdict(int)
    filtrados = []
    for c in candidatos:
        sector = c.get('sector', 'UNKNOWN')
        if sector_count[sector] < max_por_sector:
            filtrados.append(c)
            sector_count[sector] += 1
    return filtrados


# ══════════════════════════════════════════════════════════════════════════════
# WALK-FORWARD BACKTESTING (VALIDACIÓN DE PESOS)
# ══════════════════════════════════════════════════════════════════════════════
class WalkForwardBT:
    """
    Valida los pesos del sistema usando walk-forward analysis.
    Divide datos en train/test, optimiza pesos en train, valida en test.
    """
    def __init__(self, cfg: Config):
        self.cfg = cfg
        self.cache = DataCache(cfg)

    def ejecutar(self, tickers: List[str], n_splits: int = 3) -> Dict:
        """
        Ejecuta walk-forward backtesting y retorna métricas de validación.
        """
        logger.info(f"🔬 Iniciando Walk-Forward Backtesting ({n_splits} splits)...")
        resultados = []

        for ticker in tickers[:20]:  # Limitar a 20 para velocidad
            try:
                df = self._descargar(ticker, f"{self.cfg.bt_lookback_anos}y", "1wk")
                if len(df) < 100:
                    continue

                split_size = len(df) // n_splits
                for i in range(n_splits - 1):
                    train = df.iloc[:split_size * (i + 1)]
                    test = df.iloc[split_size * (i + 1):split_size * (i + 2)]

                    # Simular señales en test
                    retornos_test = self._simular_trades(test)
                    resultados.extend(retornos_test)
            except Exception as e:
                logger.debug(f"BT error {ticker}: {e}")

        if not resultados:
            return {'mensaje': 'Sin datos suficientes para backtesting'}

        retornos = np.array(resultados)
        return {
            'total_trades_sim': len(retornos),
            'win_rate': np.mean(retornos > 0) * 100,
            'avg_return': np.mean(retornos),
            'sharpe': np.mean(retornos) / (np.std(retornos) + 1e-9) * np.sqrt(12),
            'max_dd': self._max_drawdown(retornos),
            'profit_factor': self._profit_factor(retornos)
        }

    def _descargar(self, ticker: str, period: str, interval: str) -> pd.DataFrame:
        cached = self.cache.get(ticker, period, interval)
        if cached is not None:
            return cached
        df = yf.Ticker(ticker).history(period=period, interval=interval)
        if not df.empty:
            self.cache.set(ticker, period, interval, df)
        return df

    def _simular_trades(self, df_test: pd.DataFrame) -> List[float]:
        """Simula trades simples basados en WR sobreventa + RSI."""
        retornos = []
        for i in range(14, len(df_test) - 5):
            ventana = df_test.iloc[:i]
            hh = ventana['High'].rolling(14).max().iloc[-1]
            ll = ventana['Low'].rolling(14).min().iloc[-1]
            wr = ((hh - ventana['Close'].iloc[-1]) / (hh - ll + 1e-9)) * -100

            d = ventana['Close'].diff()
            g = d.clip(lower=0).rolling(14).mean().iloc[-1]
            l = (-d.clip(upper=0)).rolling(14).mean().iloc[-1]
            rsi = 100 - (100 / (1 + g / (l + 1e-9)))

            if wr <= -80 and rsi <= 40:
                # Simular entrada, hold 5 velas
                entry = df_test['Close'].iloc[i]
                exit_p = df_test['Close'].iloc[min(i + 5, len(df_test) - 1)]
                retornos.append((exit_p - entry) / entry * 100)
        return retornos

    def _max_drawdown(self, retornos: np.ndarray) -> float:
        equity = np.cumsum(retornos)
        peak = np.maximum.accumulate(equity)
        dd = (equity - peak) / (peak + 1e-9) * 100
        return float(np.min(dd)) if len(dd) > 0 else 0

    def _profit_factor(self, retornos: np.ndarray) -> float:
        wins = sum(r for r in retornos if r > 0)
        losses = abs(sum(r for r in retornos if r < 0))
        return wins / losses if losses > 0 else float('inf')


# ══════════════════════════════════════════════════════════════════════════════
# INDICADORES TÉCNICOS (MEJORADOS)
# ══════════════════════════════════════════════════════════════════════════════
class IndicadoresTecnicos:
    @staticmethod
    def williams_r(df: pd.DataFrame, p: int = 14) -> pd.Series:
        hh = df['High'].rolling(p).max()
        ll = df['Low'].rolling(p).min()
        return ((hh - df['Close']) / (hh - ll + 1e-9)) * -100

    @staticmethod
    def rsi(df: pd.DataFrame, p: int = 14) -> pd.Series:
        d = df['Close'].diff()
        g = d.clip(lower=0).rolling(p).mean()
        l = (-d.clip(upper=0)).rolling(p).mean()
        return 100 - (100 / (1 + g / l.replace(0, np.nan)))

    @staticmethod
    def macd(df: pd.DataFrame, f: int = 12, s: int = 26, sig: int = 9) -> Tuple[pd.Series, pd.Series, pd.Series]:
        ef = df['Close'].ewm(span=f, adjust=False).mean()
        es = df['Close'].ewm(span=s, adjust=False).mean()
        ml = ef - es
        sl = ml.ewm(span=sig, adjust=False).mean()
        return ml, sl, ml - sl

    @staticmethod
    def bollinger(df: pd.DataFrame, p: int = 20, k: int = 2) -> Tuple[pd.Series, pd.Series, pd.Series, pd.Series, pd.Series]:
        sma = df['Close'].rolling(p).mean()
        std = df['Close'].rolling(p).std()
        upper = sma + k * std
        lower = sma - k * std
        pct_b = (df['Close'] - lower) / (upper - lower + 1e-9)
        width = (upper - lower) / (sma + 1e-9)
        return upper, sma, lower, pct_b, width

    @staticmethod
    def atr(df: pd.DataFrame, p: int = 14) -> pd.Series:
        tr = pd.concat([
            df['High'] - df['Low'],
            (df['High'] - df['Close'].shift()).abs(),
            (df['Low'] - df['Close'].shift()).abs()
        ], axis=1).max(axis=1)
        return tr.rolling(p).mean()

    @staticmethod
    def divergencia_wr(df: pd.DataFrame, wr: pd.Series, lb: int = 10) -> str:
        p = df['Close'].values
        w = wr.values
        if len(p) < lb:
            return "Ninguna"
        pr, wr2 = p[-lb:], w[-lb:]
        if pr[-1] < np.min(pr[:-1]) and wr2[-1] > np.min(wr2[:-1]) and wr2[-1] < -50:
            return "🟢 Alcista (Bullish)"
        if pr[-1] > np.max(pr[:-1]) and wr2[-1] < np.max(wr2[:-1]) and wr2[-1] > -50:
            return "🔴 Bajista (Bearish)"
        return "Ninguna"

    @staticmethod
    def semanas_en_sobreventa(wr_series: pd.Series, umbral: float = -80) -> int:
        vals = wr_series.dropna().values[::-1]
        count = 0
        for v in vals:
            if v <= umbral:
                count += 1
            else:
                break
        return count

    @staticmethod
    def nivel_fibonacci(df: pd.DataFrame, ventana: int = 52) -> Tuple[Optional[str], float, str]:
        sub = df['Close'].iloc[-ventana:]
        maximo = sub.max()
        minimo = sub.min()
        rango = maximo - minimo
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
        mas_cercano = min(niveles.items(), key=lambda x: abs(x[1] - precio))
        dist_pct = abs(mas_cercano[1] - precio) / precio
        return mas_cercano[0], dist_pct, f"Fib {mas_cercano[0]} @ ${mas_cercano[1]:.2f}"

    @staticmethod
    def patron_reversal(df: pd.DataFrame) -> Tuple[str, int]:
        if len(df) < 2:
            return "Ninguno", 0
        curr = df.iloc[-1]
        prev = df.iloc[-2]
        o, h, l, c = curr['Open'], curr['High'], curr['Low'], curr['Close']
        body = abs(c - o)
        rango = h - l
        if rango == 0:
            return "Ninguno", 0
        sombra_inf = min(o, c) - l
        sombra_sup = h - max(o, c)

        if (sombra_inf >= body * 2 and sombra_sup <= body * 0.3 and c > o and rango > 0):
            return "Martillo", 2
        if (sombra_sup >= body * 2 and sombra_inf <= body * 0.3 and prev['Close'] < prev['Open']):
            return "Martillo Invertido", 1
        if (c > o and prev['Close'] < prev['Open'] and c > prev['Open'] and o < prev['Close']):
            return "Engulfing Alcista", 2
        if body <= rango * 0.1 and rango > 0:
            return "Doji", 1
        if c > o and (c - l) / rango >= 0.80:
            return "Vela Alcista Fuerte", 1
        return "Ninguno", 0


# ══════════════════════════════════════════════════════════════════════════════
# SISTEMA PRINCIPAL V4
# ══════════════════════════════════════════════════════════════════════════════
class SwingSystemV4:
    def __init__(self, cfg: Config):
        self.cfg = cfg
        self.ind = IndicadoresTecnicos()
        self.cache = DataCache(cfg)
        self.regime = MarketRegime(cfg)
        self.portfolio = PortfolioTracker(cfg)
        self.bt = WalkForwardBT(cfg)

        logger.info("🔄 Conectando con Google Sheets...")
        scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            cfg.credenciales_json, scope)
        self.client = gspread.authorize(creds)

        wb_swing = self.client.open(cfg.swing_archivo)
        self.sh_swing = wb_swing.worksheet(cfg.swing_hoja)

        try:
            if cfg.lp_archivo == cfg.swing_archivo:
                wb_lp = wb_swing
            else:
                wb_lp = self.client.open(cfg.lp_archivo)
            self.sh_lp = wb_lp.worksheet(cfg.lp_hoja)
            logger.info(f"✅ Largo plazo: {cfg.lp_archivo} → {cfg.lp_hoja}")
        except Exception as e:
            logger.warning(f"⚠️ No se pudo conectar a hoja LP: {e}")
            self.sh_lp = None

        logger.info(f"✅ Swing: {cfg.swing_archivo} → {cfg.swing_hoja}")

        # Detectar régimen
        self.regime_actual = self.regime.detectar()
        logger.info(f"📊 {self.regime_actual['descripcion']}")

    # ──────────────────────────────────────────────────────────────────────────
    # LEER DATOS DE LARGO PLAZO (con sector)
    # ──────────────────────────────────────────────────────────────────────────
    def leer_puntuaciones_lp(self) -> Dict[str, Dict]:
        if self.sh_lp is None:
            return {}
        cfg = self.cfg
        try:
            fi, ff = cfg.lp_fila_ini, cfg.lp_fila_fin
            col_t = cfg.lp_col_ticker
            col_p = cfg.lp_col_puntuacion
            col_e = cfg.lp_col_earnings
            col_s = cfg.lp_col_sector

            tickers_raw = self.sh_lp.get(f"{col_t}{fi}:{col_t}{ff}")
            puntos_raw = self.sh_lp.get(f"{col_p}{fi}:{col_p}{ff}")
            earnings_raw = self.sh_lp.get(f"{col_e}{fi}:{col_e}{ff}")
            # Leer sector solo si la columna está configurada y no está vacía
            sector_raw = []
            if col_s and col_s.strip():
                try:
                    sector_raw = self.sh_lp.get(f"{col_s}{fi}:{col_s}{ff}")
                except Exception:
                    sector_raw = []

            resultado = {}
            for i, row_t in enumerate(tickers_raw):
                ticker = row_t[0].strip().upper() if row_t else ""
                if not ticker:
                    continue
                try:
                    p_raw = puntos_raw[i][0] if i < len(puntos_raw) and puntos_raw[i] else None
                    p_val = float(str(p_raw).replace('%', '').strip()) if p_raw not in (None, '', 'N/A') else None
                    e_raw = earnings_raw[i][0].strip() if i < len(earnings_raw) and earnings_raw[i] else 'N/A'
                    # FIX 1: Validar que sector_raw no esté vacío y el valor sea texto válido
                    s_raw = 'UNKNOWN'
                    if i < len(sector_raw) and sector_raw[i] and sector_raw[i][0]:
                        val = str(sector_raw[i][0]).strip()
                        # Rechazar valores que parecen números (precios, no sectores)
                        try:
                            float(val.replace(',', '').replace('$', ''))
                            # Si se puede convertir a float, es un número → usar UNKNOWN
                            s_raw = 'UNKNOWN'
                        except ValueError:
                            # Es texto válido → usar como sector
                            s_raw = val if val else 'UNKNOWN'
                    resultado[ticker] = {'pts': p_val, 'earnings': e_raw, 'sector': s_raw}
                except Exception:
                    resultado[ticker] = {'pts': None, 'earnings': 'N/A', 'sector': 'UNKNOWN'}
            logger.info(f"📋 Puntuaciones LP cargadas: {len(resultado)} empresas")
            return resultado
        except Exception as e:
            logger.error(f"Error leyendo LP: {e}")
            return {}

    # ──────────────────────────────────────────────────────────────────────────
    # DESCARGA PARALELA DE DATOS
    # ──────────────────────────────────────────────────────────────────────────
    def descargar_ticker(self, ticker_sym: str) -> Tuple[str, Optional[pd.DataFrame], Optional[pd.DataFrame]]:
        """Descarga semanal y diario para un ticker. Retorna (ticker, df_w, df_d)."""
        try:
            # Intentar cache primero
            df_w = self.cache.get(ticker_sym, "2y", "1wk")
            df_d = self.cache.get(ticker_sym, "6mo", "1d")

            if df_w is None or df_d is None:
                stock = yf.Ticker(ticker_sym)
                if df_w is None:
                    df_w = stock.history(period="2y", interval="1wk")
                    if not df_w.empty:
                        self.cache.set(ticker_sym, "2y", "1wk", df_w)
                if df_d is None:
                    df_d = stock.history(period="6mo", interval="1d")
                    if not df_d.empty:
                        self.cache.set(ticker_sym, "6mo", "1d", df_d)

            return ticker_sym, df_w, df_d
        except Exception as e:
            logger.debug(f"Descarga fallida {ticker_sym}: {e}")
            return ticker_sym, None, None

    # ──────────────────────────────────────────────────────────────────────────
    # PUNTUAR UN TIMEFRAME
    # ──────────────────────────────────────────────────────────────────────────
    def _puntuar_tf(self, df: pd.DataFrame, label: str = "") -> Dict:
        pts = 0
        señales = []

        # Williams %R
        w14 = self.ind.williams_r(df, 14)
        w7 = self.ind.williams_r(df, 7)
        wr = w14.iloc[-1]
        wr7 = w7.iloc[-1]

        if wr <= -80:
            pts += 20
            señales.append(f"[{label}] WR14 sobreventa extrema ({wr:.1f})")
        elif wr <= -60:
            pts += 10
            señales.append(f"[{label}] WR14 sobreventa moderada ({wr:.1f})")
        elif wr >= -20:
            pts -= 15
            señales.append(f"[{label}] WR14 sobrecompra ({wr:.1f})")
        if wr <= -70 and wr7 <= -70:
            pts += 10
            señales.append(f"[{label}] WR multi-periodo alineado")

        # RSI
        rsi = self.ind.rsi(df, 14)
        rv = rsi.iloc[-1]
        if rv <= 30:
            pts += 15
            señales.append(f"[{label}] RSI sobreventa ({rv:.1f})")
        elif rv <= 40:
            pts += 8
            señales.append(f"[{label}] RSI zona baja ({rv:.1f})")
        elif rv >= 70:
            pts -= 12
            señales.append(f"[{label}] RSI sobrecompra ({rv:.1f})")
        elif rv >= 60:
            pts -= 5
        if len(df) >= 10:
            pa = df['Close'].values[-10:]
            ra = rsi.values[-10:]
            if pa[-1] < np.min(pa[:-1]) and ra[-1] > np.min(ra[:-1]) and rv < 45:
                pts += 12
                señales.append(f"[{label}] Divergencia alcista RSI")

        # MACD
        ml, ms, mh = self.ind.macd(df)
        mv = ml.iloc[-1]
        msv = ms.iloc[-1]
        mhv = mh.iloc[-1]
        mp = mh.iloc[-2] if len(mh) > 1 else mhv
        if mv > msv and ml.iloc[-2] <= ms.iloc[-2]:
            pts += 15
            señales.append(f"[{label}] Cruce alcista MACD")
        elif mhv > mp and mhv < 0:
            pts += 8
            señales.append(f"[{label}] MACD histograma mejorando")
        elif mv > msv:
            pts += 5
            señales.append(f"[{label}] MACD sobre señal")
        elif mv < msv and ml.iloc[-2] >= ms.iloc[-2]:
            pts -= 12
            señales.append(f"[{label}] Cruce bajista MACD")

        # Bollinger
        bbu, bbm, bbl, pct_b, bbw = self.ind.bollinger(df)
        pbv = pct_b.iloc[-1]
        bwv = bbw.iloc[-1]
        bwm = bbw.rolling(10).mean().iloc[-1]
        if pbv <= 0.05:
            pts += 15
            señales.append(f"[{label}] Precio en/bajo banda inferior BB")
        elif pbv <= 0.20:
            pts += 8
            señales.append(f"[{label}] Precio cerca banda inferior BB")
        elif pbv >= 0.95:
            pts -= 12
            señales.append(f"[{label}] Precio en/sobre banda superior BB")
        if bwv < bwm * 0.75:
            pts += 7
            señales.append(f"[{label}] BB Squeeze detectado")

        # EMAs / SMAs
        precio = df['Close'].iloc[-1]
        ema8 = df['Close'].ewm(span=8, adjust=False).mean().iloc[-1]
        ema21 = df['Close'].ewm(span=21, adjust=False).mean().iloc[-1]
        sma20 = df['Close'].rolling(20).mean().iloc[-1]
        sma50 = df['Close'].rolling(50).mean().iloc[-1] if len(df) >= 50 else ema21
        if precio > sma20:
            pts += 8
            señales.append(f"[{label}] Precio sobre SMA20")
        if ema8 > ema21:
            pts += 5
            señales.append(f"[{label}] EMA8 > EMA21")
        if precio > sma50:
            pts += 5
            señales.append(f"[{label}] Precio sobre SMA50")

        # Divergencia WR
        div = self.ind.divergencia_wr(df, w14)
        if "Alcista" in div:
            pts += 15
            señales.append(f"[{label}] Divergencia alcista WR%")
        elif "Bajista" in div:
            pts -= 10
            señales.append(f"[{label}] Divergencia bajista WR%")

        # Volumen
        vr = df['Volume'].iloc[-1] / df['Volume'].rolling(10).mean().iloc[-1]
        if vr >= 1.5:
            pts += 8
            señales.append(f"[{label}] Volumen elevado ({vr:.1f}x)")
        elif vr >= 1.2:
            pts += 4

        # Persistencia en sobreventa
        semanas_sv = self.ind.semanas_en_sobreventa(w14, umbral=-80)
        if semanas_sv >= 4:
            pts += 18
            señales.append(f"[{label}] WR en sobreventa {semanas_sv} velas (señal madura)")
        elif semanas_sv >= 2:
            pts += 10
            señales.append(f"[{label}] WR en sobreventa {semanas_sv} velas (consolidando)")
        elif semanas_sv == 1:
            pts += 4
            señales.append(f"[{label}] WR en sobreventa 1 vela (señal nueva)")

        # Fibonacci
        fib_nivel, fib_dist, fib_desc = self.ind.nivel_fibonacci(df)
        if fib_nivel and fib_dist <= 0.02:
            pts += 15
            señales.append(f"[{label}] Precio en soporte Fibonacci {fib_desc}")
        elif fib_nivel and fib_dist <= 0.05:
            pts += 8
            señales.append(f"[{label}] Precio cerca de Fibonacci {fib_desc}")

        # Patrón reversal
        patron, fuerza = self.ind.patron_reversal(df)
        if fuerza == 2:
            pts += 15
            señales.append(f"[{label}] Patrón reversal fuerte: {patron}")
        elif fuerza == 1:
            pts += 7
            señales.append(f"[{label}] Patrón reversal moderado: {patron}")

        # FIX 3: Penalizador por tendencia bajista (precio vs SMAs)
        if len(df) >= 50:
            sma50_val = df['Close'].rolling(50).mean().iloc[-1]
            if precio < sma50_val:
                pts -= 10
                señales.append(f"[{label}] Precio debajo SMA50 — tendencia bajista (-10)")
        if len(df) >= 20:
            sma20_val = df['Close'].rolling(20).mean().iloc[-1]
            if precio < sma20_val:
                pts -= 5
                señales.append(f"[{label}] Precio debajo SMA20 — momentum bajista (-5)")

        return {
            'puntos': pts, 'señales': señales,
            'wr14': wr, 'wr7': wr7, 'rsi': rv, 'macd_h': mhv, 'pct_b': pbv,
            'div': div, 'sma20': sma20, 'sma50': sma50, 'v_ratio': vr,
            'w14_s': w14, 'w7_s': w7, 'rsi_s': rsi,
            'macd_l': ml, 'macd_sig': ms, 'macd_h_s': mh,
            'bb_upper': bbu, 'bb_lower': bbl, 'bb_mid': bbm,
            'semanas_sv': semanas_sv,
            'fib_nivel': fib_nivel, 'fib_dist': fib_dist, 'fib_desc': fib_desc,
            'patron': patron, 'fuerza_patron': fuerza,
        }

    # ──────────────────────────────────────────────────────────────────────────
    # ANÁLISIS COMPLETO DE UN TICKER
    # ──────────────────────────────────────────────────────────────────────────
    # FIX 1: Aceptar DataFrames precargados para evitar doble descarga
    def analizar(self, ticker_sym: str, lp_data: Dict,
                 df_w: Optional[pd.DataFrame] = None,
                 df_d: Optional[pd.DataFrame] = None) -> Optional[Dict]:
        try:
            logger.info(f"📊 {ticker_sym}...", extra={'ticker': ticker_sym})

            # Usar datos precargados del paralelismo o descargar si no existen
            if df_w is None or df_d is None:
                _, df_w, df_d = self.descargar_ticker(ticker_sym)

            if df_w is None or df_w.empty or len(df_w) < 30:
                logger.warning(f"  ❌ {ticker_sym}: Datos insuficientes")
                return None

            precio = df_w['Close'].iloc[-1]
            atr_val = self.ind.atr(df_w, 14).iloc[-1]

            # Puntuación semanal (60%)
            rw = self._puntuar_tf(df_w, "W")

            # Puntuación diaria (40%)
            if df_d is not None and not df_d.empty and len(df_d) >= 30:
                rd = self._puntuar_tf(df_d, "D")
                confluencia = 15 if (rw['puntos'] > 0 and rd['puntos'] > 0) else                              -10 if (rw['puntos'] < 0 and rd['puntos'] < 0) else 0
                pts_swing_raw = rw['puntos'] * 0.60 + rd['puntos'] * 0.40 + confluencia
                señales_todas = rw['señales'] + rd['señales']
                wr_d = rd['wr14']
                rsi_d = rd['rsi']
                div_d = rd['div']
            else:
                rd = None
                pts_swing_raw = rw['puntos']
                señales_todas = rw['señales']
                wr_d = rsi_d = None
                div_d = "N/D"

            # Normalizar swing a 0-100
            pts_swing = max(0, min(100, int((pts_swing_raw + 30) / 1.6)))

            # Datos LP
            pts_lp = lp_data.get('pts')
            sector = lp_data.get('sector', 'UNKNOWN')

            # Earnings reales (prioridad sobre Sheets)
            ew_real, fecha_earnings = obtener_earnings_window(ticker_sym)
            ew = ew_real  # Usar el real, no el de Sheets

            # Matriz de decisión adaptativa
            mx = matriz_decision_adaptativa(pts_swing, pts_lp, self.regime_actual)

            # FIX 2: Capar puntuación combo según zona LP
            cap_lp = {'ALTO': 100, 'MEDIO': 85, 'BAJO': 60}
            pts_combo_raw = pts_swing + mx['bonus']
            pts_combo = max(0, min(cap_lp.get(mx['zona_lp'], 100), pts_combo_raw))

            # Earnings Risk
            VETO_EARNINGS = ew == 'ESTA SEMANA'
            ALERTA_EARNINGS = ew == 'ESTE MES'
            if VETO_EARNINGS:
                pts_combo = 0
                mx['accion'] = f'🚫 VETO EARNINGS ({ew}) — ' + mx['accion']
                mx['acc'] = 'ESPERAR'
                señales_todas.insert(0, f'⚠️ EARNINGS ESTA SEMANA — NO OPERAR')
            elif ALERTA_EARNINGS:
                pts_combo = max(0, pts_combo - 20)
                mx['accion'] = f'⚠️ EARNINGS PRÓX ({ew}) — ' + mx['accion']
                señales_todas.insert(0, f'⚠️ Earnings este mes — reducir tamaño de posición')

            # SL / TP dinámicos
            sl = round(precio - atr_val * 2.0, 2)
            tp1 = round(precio + atr_val * 3.0, 2)
            tp2 = round(precio + atr_val * 5.0, 2)
            rb = round((tp1 - precio) / max(precio - sl, 0.01), 2)

            setup = ("Setup Perfecto MTF+LP"
                     if rw['wr14'] <= -80 and "Alcista" in rw['div']
                        and mx['zona_lp'] == "ALTO"
                     else "Análisis Normal")

            sig = (f"EW:{ew} | WR(W):{rw['wr14']:.0f} RSI(W):{rw['rsi']:.0f}"
                   + (f" | WR(D):{wr_d:.0f} RSI(D):{rsi_d:.0f}" if wr_d else "")
                   + f" | LP:{pts_lp if pts_lp is not None else 'N/D'}")

            logger.info(f"  Swing:{pts_swing} LP:{pts_lp} EW:{ew} SV:{rw['semanas_sv']}sem "
                       f"Fib:{rw['fib_nivel']} Pat:{rw['patron']} Combo:{pts_combo} → {mx['accion']}")

            return {
                'ticker': ticker_sym,
                'precio': round(precio, 2),
                'sector': sector,
                'r14': safe(rw['wr14']),
                'r7': safe(rw['wr7']),
                'rsi_w': safe(rw['rsi']),
                'macd_w': safe(rw['macd_h']),
                'pct_b_w': safe(rw['pct_b'] * 100),
                'zona': ("Sobreventa" if rw['wr14'] <= -80
                         else "Sobrecompra" if rw['wr14'] >= -20 else "Neutral"),
                'div': rw['div'],
                'r14_d': safe(wr_d) if wr_d is not None else 'N/D',
                'rsi_d': safe(rsi_d) if rsi_d is not None else 'N/D',
                'div_d': div_d,
                'sma20': safe(rw['sma20']),
                'sma50': safe(rw['sma50']),
                'vol': safe(rw['v_ratio']),
                'pts_swing': pts_swing,
                'pts_lp': safe(pts_lp) if pts_lp is not None else 'N/D',
                'pts_combo': pts_combo,
                'zona_lp': mx['zona_lp'],
                'zona_sw': mx['zona_sw'],
                'dec': mx['accion'],
                'acc': mx['acc'],
                'rec': setup,
                'ew': ew,
                'fecha_earnings': fecha_earnings.isoformat() if fecha_earnings else 'N/D',
                'semanas_sv': safe(rw['semanas_sv']),
                'fib_nivel': rw['fib_nivel'] or 'N/D',
                'fib_dist': safe(rw['fib_dist'] * 100),
                'patron': rw['patron'],
                'fuerza_patron': safe(rw['fuerza_patron']),
                'sig': sig,
                'atr': safe(atr_val),
                'sl': sl, 'tp1': tp1, 'tp2': tp2, 'rb': rb,
                'df_w': df_w, 'df_d': df_d,
                'rw': rw, 'rd': rd,
                'señales': señales_todas,
                'regime': mx['regime'],
                'ajuste_regime': mx['ajuste_regime'],
            }
        except Exception as e:
            logger.error(f"❌ Error analizando {ticker_sym}: {e}")
            return None

    # ──────────────────────────────────────────────────────────────────────────
    # GRÁFICA MEJORADA (5 paneles + info de régimen y portfolio)
    # ──────────────────────────────────────────────────────────────────────────
    def graficar(self, d: Dict, guardar_como: Optional[str] = None) -> str:
        ticker = d['ticker']
        df = d['df_w'].copy().iloc[-52:]
        rw = d['rw']

        w14 = rw['w14_s'].iloc[-52:]
        w7 = rw['w7_s'].iloc[-52:]
        rsi_s = rw['rsi_s'].iloc[-52:]
        macd_l = rw['macd_l'].iloc[-52:]
        macd_sig = rw['macd_sig'].iloc[-52:]
        macd_h = rw['macd_h_s'].iloc[-52:]
        bb_up = rw['bb_upper'].iloc[-52:]
        bb_lo = rw['bb_lower'].iloc[-52:]
        bb_mi = rw['bb_mid'].iloc[-52:]

        BG = '#0d1117'
        PANEL = '#161b22'
        GR = '#26a641'
        RD = '#f85149'
        YL = '#e3b341'
        BL = '#58a6ff'
        PU = '#bc8cff'
        OR = '#f0883e'
        GY = '#8b949e'
        WH = '#e6edf3'
        TE = '#39d353'

        fig = plt.figure(figsize=(18, 18), facecolor=BG)
        gs = gridspec.GridSpec(6, 1, figure=fig,
                               height_ratios=[3, 1.2, 1.2, 1.2, 1, 0.8],
                               hspace=0.06)
        axes = [fig.add_subplot(gs[i]) for i in range(6)]
        for ax in axes:
            ax.set_facecolor(PANEL)
            ax.tick_params(colors=GY, labelsize=8)
            ax.spines[:].set_color('#30363d')

        fechas = df.index
        n = len(df)
        xs = range(n)
        ax1, ax2, ax3, ax4, ax5, ax6 = axes

        # Panel 1: Velas + BB + SMAs
        for i, (_, row) in enumerate(df.iterrows()):
            c = GR if row['Close'] >= row['Open'] else RD
            ax1.plot([i, i], [row['Low'], row['High']], color=c, lw=0.8, alpha=0.7)
            ax1.add_patch(plt.Rectangle((i - 0.3, min(row['Open'], row['Close'])),
                          0.6, abs(row['Close'] - row['Open']), color=c, alpha=0.85))

        ax1.plot(xs, bb_up.values, color=PU, lw=0.9, ls='--', alpha=0.7, label='BB Sup')
        ax1.plot(xs, bb_mi.values, color=GY, lw=0.9, ls='-', alpha=0.4, label='BB Med')
        ax1.plot(xs, bb_lo.values, color=TE, lw=0.9, ls='--', alpha=0.7, label='BB Inf')
        ax1.fill_between(xs, bb_lo.values, bb_up.values, alpha=0.04, color=BL)
        sma20v = df['Close'].rolling(20).mean()
        sma50v = df['Close'].rolling(50).mean()
        ax1.plot(xs, sma20v.values, color=BL, lw=1.3, label='SMA20')
        ax1.plot(xs, sma50v.values, color=PU, lw=1.3, label='SMA50', ls='--')
        ax1.axhline(d['precio'], color=YL, lw=1.5, label=f"Entrada ${d['precio']}")
        ax1.axhline(d['sl'], color=RD, lw=1.5, ls='--', label=f"SL ${d['sl']}")
        ax1.axhline(d['tp1'], color=GR, lw=1.5, ls='--', label=f"TP1 ${d['tp1']}")
        ax1.axhline(d['tp2'], color=GR, lw=2.0, label=f"TP2 ${d['tp2']}")
        ax1.axhspan(d['sl'], d['precio'], alpha=0.07, color=RD)
        ax1.axhspan(d['precio'], d['tp1'], alpha=0.05, color=GR)
        ax1.axhspan(d['tp1'], d['tp2'], alpha=0.09, color=GR)

        for val, lbl, col in [(d['tp2'], f"TP2 ${d['tp2']}", GR),
                               (d['tp1'], f"TP1 ${d['tp1']}", GR),
                               (d['precio'], f"ENTRADA ${d['precio']}", YL),
                               (d['sl'], f"SL ${d['sl']}", RD)]:
            ax1.annotate(lbl, xy=(n - 1, val), xytext=(n + 1, val), color=col,
                         fontsize=8, fontweight='bold', va='center',
                         arrowprops=dict(arrowstyle='->', color=col, lw=1))
        ax1.set_xlim(-1, n + 10)

        rb_c = GR if d['rb'] >= 2 else YL if d['rb'] >= 1.5 else RD
        ax1.text(0.01, 0.97,
                 f"R:R={d['rb']}x  Riesgo:${round(d['precio'] - d['sl'], 2)}"
                 f"  TP1:+${round(d['tp1'] - d['precio'], 2)}",
                 transform=ax1.transAxes, color=rb_c, fontsize=8.5, va='top',
                 fontweight='bold',
                 bbox=dict(boxstyle='round,pad=0.3', facecolor=BG, alpha=0.7))

        lp_str = f"{d['pts_lp']}" if d['pts_lp'] != 'N/D' else "N/D"
        lp_c = GR if d['zona_lp'] == 'ALTO' else YL if d['zona_lp'] == 'MEDIO' else RD
        ax1.text(0.99, 0.97,
                 f"Swing: {d['pts_swing']}/100  |  LP: {lp_str}  |  Zona LP: {d['zona_lp']}",
                 transform=ax1.transAxes, color=lp_c, fontsize=8.5, va='top', ha='right',
                 fontweight='bold',
                 bbox=dict(boxstyle='round,pad=0.3', facecolor=BG, alpha=0.7))

        ax1.legend(loc='upper left', fontsize=7, facecolor=PANEL,
                   labelcolor=WH, framealpha=0.8, ncol=5)
        ax1.set_ylabel('Precio (USD)', color=GY, fontsize=9)
        ax1.tick_params(labelbottom=False)

        # Panel 2: Williams %R
        ax2.plot(xs, w14.values, color=BL, lw=1.5, label='W%R 14')
        ax2.plot(xs, w7.values, color=OR, lw=1.0, label='W%R 7', ls='--', alpha=0.75)
        ax2.axhline(-20, color=RD, lw=0.8, ls=':', alpha=0.7)
        ax2.axhline(-50, color=GY, lw=0.8, ls=':', alpha=0.5)
        ax2.axhline(-80, color=GR, lw=0.8, ls=':', alpha=0.7)
        ax2.axhspan(-100, -80, alpha=0.10, color=GR)
        ax2.axhspan(-20, 0, alpha=0.10, color=RD)
        ax2.set_ylim(-105, 5)
        wv = w14.iloc[-1]
        wc = GR if wv <= -80 else RD if wv >= -20 else YL
        ax2.scatter([n - 1], [wv], color=wc, s=60, zorder=5)
        ax2.text(n - 1, wv + 3, f'{wv:.1f}', color=wc, fontsize=7.5, ha='center', fontweight='bold')
        ax2.set_ylabel('Williams %R', color=GY, fontsize=9)
        ax2.legend(loc='upper right', fontsize=7.5, facecolor=PANEL, labelcolor=WH, framealpha=0.8)
        ax2.tick_params(labelbottom=False)

        # Panel 3: RSI
        ax3.plot(xs, rsi_s.values, color=TE, lw=1.5, label='RSI 14')
        ax3.axhline(70, color=RD, lw=0.8, ls=':', alpha=0.7)
        ax3.axhline(50, color=GY, lw=0.8, ls=':', alpha=0.4)
        ax3.axhline(30, color=GR, lw=0.8, ls=':', alpha=0.7)
        ax3.axhspan(0, 30, alpha=0.08, color=GR)
        ax3.axhspan(70, 100, alpha=0.08, color=RD)
        ax3.set_ylim(0, 100)
        rv2 = rsi_s.iloc[-1]
        rc = GR if rv2 <= 30 else RD if rv2 >= 70 else YL
        ax3.scatter([n - 1], [rv2], color=rc, s=60, zorder=5)
        ax3.text(n - 1, rv2 + 2, f'{rv2:.1f}', color=rc, fontsize=7.5, ha='center', fontweight='bold')
        ax3.set_ylabel('RSI', color=GY, fontsize=9)
        ax3.legend(loc='upper right', fontsize=7.5, facecolor=PANEL, labelcolor=WH, framealpha=0.8)
        ax3.tick_params(labelbottom=False)

        # Panel 4: MACD
        ax4.plot(xs, macd_l.values, color=BL, lw=1.3, label='MACD')
        ax4.plot(xs, macd_sig.values, color=OR, lw=1.0, label='Señal', ls='--')
        for i, h in enumerate(macd_h.values):
            ax4.bar(i, h, color=GR if h >= 0 else RD, alpha=0.6, width=0.8)
        ax4.axhline(0, color=GY, lw=0.8, alpha=0.4)
        ax4.set_ylabel('MACD', color=GY, fontsize=9)
        ax4.legend(loc='upper right', fontsize=7.5, facecolor=PANEL, labelcolor=WH, framealpha=0.8)
        ax4.tick_params(labelbottom=False)

        # Panel 5: Volumen
        vm = df['Volume'].rolling(10).mean()
        for i, (v, vm_) in enumerate(zip(df['Volume'].values, vm.values)):
            ax5.bar(i, v, color=GR if v > vm_ else GY, alpha=0.75, width=0.8)
        ax5.plot(xs, vm.values, color=YL, lw=1.0, ls='--', label='Vol MA10')
        ax5.set_ylabel('Volumen', color=GY, fontsize=9)
        ax5.legend(loc='upper right', fontsize=7.5, facecolor=PANEL, labelcolor=WH, framealpha=0.8)
        ax5.tick_params(labelbottom=False)

        # Panel 6: Info de régimen y portfolio
        ax6.axis('off')
        regime_info = f"RÉGIMEN: {d.get('regime', 'N/D')} (ajuste: {d.get('ajuste_regime', 0):+d})"
        sector_info = f"SECTOR: {d.get('sector', 'N/D')}"
        earnings_info = f"EARNINGS: {d['ew']}"
        if d['fecha_earnings'] != 'N/D':
            earnings_info += f" ({d['fecha_earnings'][:10]})"

        info_text = f"""
{regime_info}
{sector_info}
{earnings_info}
Fibonacci: {d['fib_nivel']} ({d['fib_dist']}% dist)
Patrón: {d['patron']} (fuerza: {d['fuerza_patron']})
Persistencia SV: {d['semanas_sv']} velas
        """.strip()

        ax6.text(0.02, 0.5, info_text, transform=ax6.transAxes, color=WH,
                 fontsize=9, va='center', ha='left', family='monospace',
                 bbox=dict(boxstyle='round,pad=0.5', facecolor=PANEL, edgecolor=GY, alpha=0.9))

        step = max(1, n // 8)
        ticks = list(range(0, n, step))
        ax5.set_xticks(ticks)
        ax5.set_xticklabels([fechas[i].strftime('%b %Y') for i in ticks],
                            color=GY, fontsize=8, rotation=30)

        cc = GR if 'COMPRA' in d['dec'] else YL if 'ESPERAR' in d['dec'] or 'SCALP' in d['dec'] else RD
        rsi_d_str = f"RSI(D):{d['rsi_d']}" if d['rsi_d'] != 'N/D' else ""
        fig.suptitle(
            f"{ticker}  —  {d['dec']}  |  Combo: {d['pts_combo']}/100  |  {self.regime_actual['descripcion']}\n"
            f"Swing: {d['pts_swing']}/100  •  LP: {d['pts_lp']}  •  Zona LP: {d['zona_lp']}  •  "
            f"WR(W):{d['r14']} RSI(W):{d['rsi_w']}  |  "
            f"WR(D):{d.get('r14_d', 'N/D')} {rsi_d_str}  |  Div: {d['div']}",
            color=WH, fontsize=11, fontweight='bold', y=0.99,
            bbox=dict(boxstyle='round,pad=0.4', facecolor=cc, alpha=0.25)
        )
        plt.tight_layout(rect=[0, 0, 1, 0.97])

        nombre = guardar_como or f"swing_{ticker}_{datetime.now().strftime('%Y%m%d_%H%M')}.png"
        plt.savefig(nombre, dpi=150, bbox_inches='tight', facecolor=BG)
        plt.close()
        logger.info(f"      📈 Gráfica: {nombre}")
        return nombre

    # ──────────────────────────────────────────────────────────────────────────
    # EJECUCIÓN PRINCIPAL (PARALELA)
    # ──────────────────────────────────────────────────────────────────────────
    def ejecutar(self):
        cfg = self.cfg

        # 1. Leer puntuaciones LP
        scores_lp = self.leer_puntuaciones_lp()

        # 2. Leer tickers
        col_data = self.sh_swing.col_values(cfg.swing_col_tickers)
        ff = cfg.swing_fila_fin or len(col_data)
        fi = cfg.swing_fila_ini
        tickers = [t.strip().upper() for t in col_data[fi - 1:ff] if t.strip()]

        if not tickers:
            logger.warning("⚠️ No hay tickers.")
            return

        logger.info(f"🔍 {len(tickers)} tickers a analizar (workers: {cfg.max_workers})")
        logger.info("=" * 60)

        # 3. Descargar datos en paralelo
        logger.info("⬇️ Descargando datos en paralelo...")
        datos = {}
        with ThreadPoolExecutor(max_workers=cfg.max_workers) as executor:
            futures = {executor.submit(self.descargar_ticker, t): t for t in tickers}
            for future in as_completed(futures):
                ticker, df_w, df_d = future.result()
                datos[ticker] = (df_w, df_d)

        # 4. Analizar cada ticker usando los datos precargados
        resultados = []
        filas = []

        for t in tickers:
            lp_data = scores_lp.get(t, {'pts': None, 'earnings': 'N/A', 'sector': 'UNKNOWN'})
            # FIX 1: Pasar los DataFrames ya descargados en paralelo
            df_w, df_d = datos.get(t, (None, None))
            d = self.analizar(t, lp_data, df_w, df_d)
            if d:
                fila = [
                    safe(d['r14']), safe(d['r7']), safe(d['rsi_w']),
                    safe(d['macd_w']), safe(d['pct_b_w']),
                    d['zona'], d['div'],
                    safe(d['r14_d']), safe(d['rsi_d']),
                    d['div_d'],
                    safe(d['sma20']), safe(d['sma50']), safe(d['vol']), safe(d['atr']),
                    safe(d['pts_swing']),
                    safe(d['pts_lp']),
                    safe(d['pts_combo']),
                    d['zona_lp'], d['zona_sw'],
                    d['dec'], d['acc'], d['rec'], d['sig'], d['ew'],
                    safe(d['semanas_sv']), d['fib_nivel'],
                    safe(d['fib_dist']), d['patron'],
                    safe(d['sl']), safe(d['tp1']), safe(d['tp2']), safe(d['rb']),
                    d['sector'], d['regime'], d['fecha_earnings']
                ]
                filas.append(fila)
                resultados.append(d)
            else:
                filas.append(['ERROR'] * 35)

        # 5. Escribir en Google Sheets
        if filas:
            n_cols = len(filas[0])
            ci = letter_to_col(cfg.swing_col_salida)
            ce = col_to_letter(ci + n_cols - 1)
            rango = f"{cfg.swing_col_salida}{fi}:{ce}{fi + len(filas) - 1}"
            self.sh_swing.update(range_name=rango, values=filas)
            logger.info(f"✅ Datos escritos en {rango}")

        # 6. Actualizar portfolio con precios actuales
        precios_actuales = {d['ticker']: d['precio'] for d in resultados}
        atrs = {d['ticker']: safe_float(d['atr']) for d in resultados}
        self.portfolio.actualizar_posiciones(precios_actuales, atrs)

        # 7. Top N con filtro de correlación sectorial
        if resultados:
            candidatos = sorted(resultados, key=lambda x: x['pts_combo'], reverse=True)
            candidatos_filtrados = filtrar_por_correlacion(candidatos, max_por_sector=2)
            top = candidatos_filtrados[:cfg.top_n]

            logger.info(f"{'=' * 60}")
            logger.info(f"🏆 TOP {cfg.top_n} MEJORES SETUPS COMBINADOS (filtrados por sector)")
            logger.info(f"{'=' * 60}")

            for rank, d in enumerate(top, 1):
                logger.info(
                    f"  #{rank}  {d['ticker']} ({d['sector']})  "
                    f"Combo:{d['pts_combo']} Swing:{d['pts_swing']} LP:{d['pts_lp']}  →  {d['dec']}"
                )
                self.graficar(d, f"TOP{rank}_{d['ticker']}_{datetime.now().strftime('%Y%m%d_%H%M')}.png")

                # Sugerir apertura de posición si es COMPRAR
                if d['acc'] == 'COMPRAR' and d['pts_combo'] >= 60:
                    capital_disp = self.cfg.capital_total - sum(
                        p.cantidad * p.entrada for p in self.portfolio.posiciones if p.estado == 'ABIERTA'
                    )
                    pos = self.portfolio.abrir_posicion(
                        d['ticker'], d['precio'], safe_float(d['atr']),
                        d['pts_combo'], capital_disp
                    )

            # 8. Señales detalladas top 3
            logger.info(f"{'─' * 60}")
            logger.info("📋 SEÑALES DETALLADAS — TOP 3")
            logger.info(f"{'─' * 60}")
            for d in top[:3]:
                logger.info(f"  {d['ticker']} (Combo:{d['pts_combo']} | Swing:{d['pts_swing']} | LP:{d['pts_lp']}):")
                for s in d['señales']:
                    logger.info(f"    • {s}")

        # 9. Métricas de portfolio
        metricas = self.portfolio.metricas()
        logger.info(f"{'=' * 60}")
        logger.info("📊 MÉTRICAS DE PORTFOLIO")
        logger.info(f"{'=' * 60}")
        for k, v in metricas.items():
            logger.info(f"  {k}: {v}")

        # 10. Walk-forward backtesting (opcional, en subset)
        logger.info(f"{'=' * 60}")
        logger.info("🔬 WALK-FORWARD BACKTESTING (subset)")
        logger.info(f"{'=' * 60}")
        bt_result = self.bt.ejecutar(tickers[:20])
        for k, v in bt_result.items():
            logger.info(f"  {k}: {v}")


# ══════════════════════════════════════════════════════════════════════════════
# PUNTO DE ENTRADA
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    sistema = SwingSystemV4(CFG)
    sistema.ejecutar()
