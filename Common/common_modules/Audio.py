# Audio.py  — A/LAeq & LAF metrics with alignment tools and soft-cal
from __future__ import annotations
import argparse, sys, time, json, csv
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Dict, Tuple

import numpy as np
import sounddevice as sd
from scipy.signal import sosfilt, bilinear_zpk, zpk2sos, sosfreqz


# ---------------- A- and C-weighting (digital, unity @ 1 kHz) ----------------
def _normalize_sos_at_1k(sos: np.ndarray, fs: int) -> np.ndarray:
    w0 = 2*np.pi*1000.0 / fs
    _, h = sosfreqz(sos, worN=[w0])
    g = 1.0 / np.abs(h[0])
    sos = sos.copy()
    sos[0, :3] *= g
    return sos

def a_weighting_sos(fs: int) -> np.ndarray:
    f1, f2, f3, f4 = 20.598997, 107.65265, 737.86223, 12194.217
    w1, w2, w3, w4 = 2*np.pi*np.array([f1, f2, f3, f4], dtype=np.float64)
    z = [0.0, 0.0]
    p = [-w1, -w2, -w3, -w4]
    zd, pd, kd = bilinear_zpk(z, p, 1.0, fs=fs)
    sos = zpk2sos(zd, pd, kd).astype(np.float64)
    return _normalize_sos_at_1k(sos, fs)

def c_weighting_sos(fs: int) -> np.ndarray:
    f1, f4 = 20.598997, 12194.217
    w1, w4 = 2*np.pi*np.array([f1, f4], dtype=np.float64)
    z = [0.0, 0.0]                       # double zero at DC
    p = [1j*w1, -1j*w1, 1j*w4, -1j*w4]   # complex poles
    zd, pd, kd = bilinear_zpk(z, p, 1.0, fs=fs)
    sos = zpk2sos(zd, pd, kd).astype(np.float64)
    return _normalize_sos_at_1k(sos, fs)


# ---------------- Config ----------------
@dataclass
class CaptureConfig:
    device_name_hint: Optional[str] = None   # e.g. "UR22", "ASIO", "WASAPI"
    samplerate: int = 48000
    # open N channels from device. If 2, we can auto-pick louder channel.
    open_channels: int = 2
    # analysis window control (to align with ARTA)
    duration_s: float = 10.0                 # total capture length
    start_offset_ms: int = 0                 # skip this much at start before analysis
    window_sec: Optional[float] = None       # if None, use [start_offset .. end]
    pre_roll_ms: int = 200                   # pre-roll before first read (helps catch early bursts)
    auto_channel: bool = True                # pick louder of L/R when open_channels>=2
    average_lr: bool = False                 # if True and open_channels>=2, average L/R
    latency: str = "low"
    # calibration
    dbfs_to_dbspl: Optional[float] = None
    cal_file: Path = Path("audio_cal.json")
    # extras
    dump_trace_csv: Optional[Path] = None    # write LAF(t) dB trace here
    debug: bool = False


# ---------------- Device helpers ----------------
def find_input_device_id(name_hint: Optional[str]) -> Optional[int]:
    if not name_hint:
        return None
    name_hint = name_hint.lower()
    devs = sd.query_devices()
    apis = sd.query_hostapis()
    for i, d in enumerate(devs):
        if d.get("max_input_channels", 0) <= 0:
            continue
        api_name = apis[d["hostapi"]]["name"]
        if name_hint in d["name"].lower() or name_hint in api_name.lower():
            return i
    return None

def device_name_from_id(dev_id: Optional[int]) -> str:
    devs = sd.query_devices()
    if isinstance(dev_id, int):
        return devs[dev_id]["name"]
    try:
        default_in = sd.default.device[0]
        return devs[default_in]["name"] if isinstance(default_in, int) else "SYSTEM DEFAULT"
    except Exception:
        return "UNKNOWN"


# ---------------- Calibration store ----------------
def _cal_key(device_name: str, cfg: CaptureConfig) -> str:
    return f"{device_name}|fs={cfg.samplerate}|opench={cfg.open_channels}"

def load_calibration(cfg: CaptureConfig, device_name: str) -> Optional[float]:
    if not cfg.cal_file.exists():
        return None
    try:
        data = json.loads(cfg.cal_file.read_text())
    except Exception:
        return None
    return data.get(_cal_key(device_name, cfg))

def save_calibration(cfg: CaptureConfig, device_name: str, offset: float) -> None:
    data = {}
    if cfg.cal_file.exists():
        try:
            data = json.loads(cfg.cal_file.read_text())
        except Exception:
            data = {}
    data[_cal_key(device_name, cfg)] = float(offset)
    cfg.cal_file.write_text(json.dumps(data, indent=2))


# ---------------- Capture ----------------
def _pick_channel(arr: np.ndarray, auto: bool, average: bool) -> np.ndarray:
    if arr.ndim == 1 or arr.shape[1] == 1:
        return arr.reshape(-1).astype(np.float64)
    if average:
        return arr.mean(axis=1).astype(np.float64)
    if auto:
        rms = np.sqrt(np.mean(arr.astype(np.float64)**2, axis=0) + 1e-30)
        idx = int(np.argmax(rms))
        return arr[:, idx].astype(np.float64)
    # default to left
    return arr[:, 0].astype(np.float64)

def record_raw(cfg: CaptureConfig, device_id: Optional[int]) -> np.ndarray:
    frames = int(cfg.duration_s * cfg.samplerate)
    ch_to_open = max(1, cfg.open_channels)

    sd.check_input_settings(device=device_id, channels=ch_to_open, samplerate=cfg.samplerate)

    with sd.InputStream(device=device_id,
                        channels=ch_to_open,
                        samplerate=cfg.samplerate,
                        dtype="float32",
                        latency=cfg.latency) as stream:
        if cfg.pre_roll_ms > 0:
            time.sleep(cfg.pre_roll_ms / 1000.0)
        data, _ = stream.read(frames)  # (frames, ch)

    x = _pick_channel(data, auto=cfg.auto_channel, average=cfg.average_lr)
    if cfg.debug:
        raw_rms = float(np.sqrt(np.mean(x**2) + 1e-30))
        raw_pk  = float(np.max(np.abs(x)) + 1e-30)
        print(f"[debug] raw: rmsFS={raw_rms:.3e}, peakFS={raw_pk:.3e}")
    return x


# ---------------- Detectors & metrics ----------------
def _laf_fast_env_sq(x: np.ndarray, fs: int, tau: float = 0.125) -> np.ndarray:
    """Exponential IEC 'Fast' detector operating on mean-square."""
    alpha = 1.0 - np.exp(-1.0/(tau*fs))
    ms = x.astype(np.float64)**2
    y = 0.0
    env = np.empty_like(ms)
    for i, v in enumerate(ms):
        y += alpha*(v - y)
        env[i] = y
    return env  # mean-square

def _subwindow(x: np.ndarray, fs: int, start_ms: int, window_sec: Optional[float]) -> np.ndarray:
    i0 = max(int(start_ms/1000 * fs), 0)
    if window_sec is None:
        return x[i0:]
    i1 = min(i0 + int(window_sec * fs), len(x))
    return x[i0:i1]

def compute_metrics(x: np.ndarray, fs: int, dbfs_to_dbspl: float,
                    start_offset_ms: int = 0, window_sec: Optional[float] = None,
                    dump_trace_csv: Optional[Path] = None) -> Dict[str, float]:
    # A-weight, unity @ 1k
    sosA = a_weighting_sos(fs)
    xA = sosfilt(sosA, x)

    xa = _subwindow(xA, fs, start_offset_ms, window_sec)

    # LAeq over window
    rmsA = float(np.sqrt(np.mean(xa**2) + 1e-30))
    leqA = 20*np.log10(rmsA) + dbfs_to_dbspl

    # LAF (exponential τ=125 ms) — ARTA-like
    env_ms = _laf_fast_env_sq(xa, fs, tau=0.125)
    laf_series_db = 10*np.log10(env_ms + 1e-30) + dbfs_to_dbspl
    lafmax = float(np.max(laf_series_db))
    lafmin = float(np.min(laf_series_db))

    # A-weighted sample peak (not LCpk)
    la_peak = 20*np.log10(float(np.max(np.abs(xa)) + 1e-30)) + dbfs_to_dbspl

    if dump_trace_csv is not None:
        t0 = start_offset_ms/1000.0
        with open(dump_trace_csv, "w", newline="") as f:
            w = csv.writer(f)
            w.writerow(["t_sec", "LAF_dB"])
            for i, val in enumerate(laf_series_db):
                w.writerow([t0 + i/fs, float(val)])

    return {"LeqA": leqA, "LAFmax": lafmax, "LAFmin": lafmin, "LApeak": la_peak}


# ---------------- Public API ----------------
def measure_once(cfg: CaptureConfig) -> Dict[str, object]:
    dev_id = find_input_device_id(cfg.device_name_hint)
    dev_name = device_name_from_id(dev_id)

    # resolve calibration
    if cfg.dbfs_to_dbspl is None:
        loaded = load_calibration(cfg, dev_name)
        if loaded is None:
            if cfg.debug:
                print("[debug] No calibration found; using placeholder 94.0 dB offset")
            cfg.dbfs_to_dbspl = 94.0
        else:
            cfg.dbfs_to_dbspl = float(loaded)

    raw = record_raw(cfg, dev_id)
    metrics = compute_metrics(raw, cfg.samplerate, cfg.dbfs_to_dbspl,
                              start_offset_ms=cfg.start_offset_ms,
                              window_sec=cfg.window_sec,
                              dump_trace_csv=cfg.dump_trace_csv)
    return {"device_id": dev_id, "device_name": dev_name, "raw": raw, "fs": cfg.samplerate, "metrics": metrics}


# ---------------- Calibration (hard + soft) ----------------
def _tone_snr_db(x: np.ndarray, fs: int, f0: float = 1000.0) -> float:
    n = len(x)
    if n < 2048:
        return -np.inf
    win = np.hanning(n)
    X = np.fft.rfft(x * win)
    freqs = np.fft.rfftfreq(n, 1/fs)
    k0 = int(np.argmin(np.abs(freqs - f0)))
    sig_pow = (np.abs(X[k0])**2)
    nb = 5
    noise = np.copy(np.abs(X)**2)
    i0 = max(k0-nb, 1); i1 = min(k0+nb+1, len(noise))
    noise[i0:i1] = 0; noise[0:1] = 0
    noise_pow = noise.mean() + 1e-30
    return 10*np.log10(sig_pow / noise_pow + 1e-30)

def calibrate(cfg: CaptureConfig, known_spl_db: float = 94.0) -> Tuple[float, float]:
    dev_id = find_input_device_id(cfg.device_name_hint)
    dev_name = device_name_from_id(dev_id)
    x = record_raw(cfg, dev_id)
    snr_db = _tone_snr_db(x, cfg.samplerate, 1000.0)
    if snr_db < 20:
        raise RuntimeError(f"Calibration aborted: 1 kHz tone not detected (SNR {snr_db:.1f} dB).")
    xa = sosfilt(a_weighting_sos(cfg.samplerate), x)
    xa = _subwindow(xa, cfg.samplerate, cfg.start_offset_ms, cfg.window_sec)
    rms_fs = float(np.sqrt(np.mean(xa**2) + 1e-30))
    offset = known_spl_db - 20*np.log10(rms_fs)
    if not (80.0 <= offset <= 150.0):
        raise RuntimeError(f"Calibration aborted: computed offset {offset:.2f} dB looks wrong.")
    save_calibration(cfg, dev_name, offset)
    cfg.dbfs_to_dbspl = offset
    return offset, rms_fs

def calibrate_soft(cfg: CaptureConfig, reference_spl_db: float) -> float:
    dev_id = find_input_device_id(cfg.device_name_hint)
    dev_name = device_name_from_id(dev_id)
    x = record_raw(cfg, dev_id)
    xa = sosfilt(a_weighting_sos(cfg.samplerate), x)
    xa = _subwindow(xa, cfg.samplerate, cfg.start_offset_ms, cfg.window_sec)
    rms_fs = float(np.sqrt(np.mean(xa**2) + 1e-30))
    offset = reference_spl_db - 20*np.log10(rms_fs)
    if not (80.0 <= offset <= 150.0):
        raise RuntimeError(f"Soft calibration aborted: computed offset {offset:.2f} dB looks wrong.")
    save_calibration(cfg, dev_name, offset)
    cfg.dbfs_to_dbspl = offset
    return offset


# ---------------- Self-check (DSP only) ----------------
def selfcheck_dsp(fs: int) -> bool:
    sos = a_weighting_sos(fs)
    w0 = 2*np.pi*1000.0 / fs
    _, h = sosfreqz(sos, worN=[w0])
    gain_db = 20*np.log10(np.abs(h[0]))
    ok_gain = abs(gain_db) < 0.1
    t = np.arange(int(fs*1.0))/fs
    A = 0.5
    y = sosfilt(sos, np.sin(2*np.pi*1000.0*t)*A)
    y = y[int(0.2*fs):]
    rms = float(np.sqrt(np.mean(y**2) + 1e-30))
    expected = A/np.sqrt(2.0)
    err_db = 20*np.log10(rms/expected + 1e-30)
    print(f"[selfcheck] A-weight @1k gain = {gain_db:+.02f} dB; 1k sine RMS err = {err_db:+.02f} dB")
    return ok_gain and abs(err_db) < 0.2


__all__ = [
    "CaptureConfig",
    "measure_once",
    "calibrate",
    "calibrate_soft",
    "load_calibration",
    "save_calibration",
    "selfcheck_dsp",
]


# ---------------- CLI ----------------
if __name__ == "__main__":
    p = argparse.ArgumentParser(description="A-weighted SPL metrics with ARTA-like 'Fast' and alignment tools")
    p.add_argument("--device", default="UR22", help="Device name hint")
    p.add_argument("--fs", type=int, default=48000)
    p.add_argument("--open-ch", type=int, default=2, help="Channels to open from device (1 or 2)")
    p.add_argument("--dur", type=float, default=10.0, help="Total capture length (s)")
    p.add_argument("--start-offset-ms", type=int, default=0, help="Start offset within capture for analysis")
    p.add_argument("--window-sec", type=float, default=None, help="Analysis window length (s); default = to end")
    p.add_argument("--no-auto-channel", action="store_true", help="Disable auto-pick louder channel")
    p.add_argument("--avg-lr", action="store_true", help="Average L/R instead of picking louder one")
    p.add_argument("--pre-roll-ms", type=int, default=200)
    p.add_argument("--debug", action="store_true")
    p.add_argument("--soft-cal", type=float, help="Soft-calibrate to this LAeq (e.g. ARTA LAeq)")
    p.add_argument("--calibrate", action="store_true", help="Hard cal with 1 kHz @ known SPL on mic")
    p.add_argument("--known-spl", type=float, default=94.0)
    p.add_argument("--trace", type=Path, help="Write LAF time series (t, LAF_dB) to CSV")
    args = p.parse_args()

    selfcheck_dsp(args.fs)

    cfg = CaptureConfig(
        device_name_hint=args.device,
        samplerate=args.fs,
        open_channels=args.open_ch,
        duration_s=args.dur,
        start_offset_ms=args.start_offset_ms,
        window_sec=args.window_sec if args.window_sec is not None else None,
        pre_roll_ms=args.pre_roll_ms,
        auto_channel=not args.no_auto_channel,
        average_lr=args.avg_lr,
        dump_trace_csv=args.trace,
        debug=args.debug
    )

    if args.soft_cal is not None:
        off = calibrate_soft(cfg, args.soft_cal)
        print(f"[cal] Soft-cal offset saved: {off:.2f} dB -> {cfg.cal_file}")

    if args.calibrate:
        off, rms = calibrate(cfg, known_spl_db=args.known_spl)
        print(f"[cal] Hard-cal offset saved: {off:.2f} dB (rmsFS={rms:.3e}) -> {cfg.cal_file}")

    res = measure_once(cfg)
    print("Device:", res["device_name"])
    print(res["metrics"])
