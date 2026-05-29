"""
terrabrasilis_to_geoparquet.py
------------------------------
Processa ZIPs do TerraBrasilis:

  VETORES (SHP / GPKG):
    reproj ESRI:102033 -> filtro Polygon/MultiPolygon -> repair_geometry
    -> grid 25x25 km -> tile_id nos atributos -> GeoParquet por tile

  RASTERS (TIF / TIFF):
    reproj ESRI:102033 -> corte em tiles 25x25 km alinhados ao mesmo grid
    -> GeoTIFF Cloud-Optimized (COG) por tile, otimizado para zonal stats

Saida:
  C:\\Amintas\\Prodes\\zip\\<hoje>\\geoparquet\\   <- vetores
  C:\\Amintas\\Prodes\\zip\\<hoje>\\raster\\        <- rasters

Dependencias instaladas automaticamente.
"""

import sys
import subprocess


# ---------------------------------------------------------------------------
# Auto-instalacao
# ---------------------------------------------------------------------------

def _instalar_pacotes(pacotes):
    print("[setup] Instalando: " + ", ".join(pacotes) + " ...")
    estrategias = [
        ["uv", "pip", "install", "--quiet"] + pacotes,
        [sys.executable, "-m", "pip", "install", "--quiet"] + pacotes,
        [sys.executable, "-m", "pip", "install", "--quiet",
         "--break-system-packages"] + pacotes,
    ]
    for cmd in estrategias:
        try:
            subprocess.check_call(cmd, stderr=subprocess.DEVNULL)
            print("[setup] OK.\n")
            return
        except (subprocess.CalledProcessError, FileNotFoundError):
            continue
    print("[ERRO] Falha ao instalar: " + " ".join(pacotes))
    sys.exit(1)


def _garantir_deps():
    faltando = []
    for pkg, mod in [
        ("geopandas",  "geopandas"),
        ("pyarrow",    "pyarrow"),
        ("shapely",    "shapely"),
        ("rasterio",   "rasterio"),
        ("tqdm",       "tqdm"),
    ]:
        try:
            __import__(mod)
        except ImportError:
            faltando.append(pkg)
    if faltando:
        _instalar_pacotes(faltando)


_garantir_deps()


import zipfile
import tempfile
import shutil
import warnings
import logging
import json
import math
import pandas as pd
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

import numpy as np
import geopandas as gpd
import rasterio
import rasterio.warp
import rasterio.mask
import rasterio.transform
from rasterio.crs import CRS
from rasterio.enums import Resampling
from rasterio.warp import calculate_default_transform, reproject
from shapely.geometry import box
from shapely.validation import make_valid
from tqdm import tqdm

warnings.filterwarnings("ignore")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Configuracoes
# ---------------------------------------------------------------------------

DATA_HOJE     = datetime.now().strftime("%Y-%m-%d")
PASTA_ZIPS    = Path(r"C:\Amintas\Prodes\zip") / DATA_HOJE
PASTA_VET     = PASTA_ZIPS / "geoparquet"
PASTA_RAS     = PASTA_ZIPS / "raster"
CRS_ALVO      = "ESRI:102033"
CRS_RASTERIO  = CRS.from_user_input(CRS_ALVO)
TILE_M        = 25_000      # 25 km em metros
MAX_WORKERS   = 4
TIPOS_VALIDOS = {"Polygon", "MultiPolygon"}

# Raster: tamanho do tile interno COG (pixels) - otimizado para zonal stats
# 512x512 eh o padrao COG; com 64 GB de RAM podemos usar janelas grandes
COG_TILE_PX   = 512         # tile interno do COG
RAM_GB        = 64
# Quantos tiles 25km processar simultaneamente em memoria (raster)
# Cada tile 25km a 30m resolucao tem ~833x833 px ~ 5 MB float32 -> cabe muito
MAX_RAM_TILES = 32


# ---------------------------------------------------------------------------
# Grid helper (compartilhado vetor + raster)
# ---------------------------------------------------------------------------

def bbox_para_tiles(minx, miny, maxx, maxy, tamanho=TILE_M):
    """
    Gera lista de (tile_id, x0, y0, x1, y1) alinhados ao grid global.
    O grid e ancorado em multiplos de `tamanho` para garantir alinhamento
    entre camadas diferentes.
    """
    ox = (minx // tamanho) * tamanho
    oy = (miny // tamanho) * tamanho
    tiles = []
    col = 0
    x = ox
    while x < maxx:
        row = 0
        y = oy
        while y < maxy:
            tile_id = "r{:04d}_c{:04d}".format(row, col)
            tiles.append((tile_id, x, y, x + tamanho, y + tamanho))
            y += tamanho
            row += 1
        x += tamanho
        col += 1
    return tiles


# ---------------------------------------------------------------------------
# VETORES
# ---------------------------------------------------------------------------

def reparar_geometria(gdf):
    def _fix(geom):
        if geom is None or geom.is_empty:
            return None
        try:
            g = geom.buffer(0)
            if not g.is_valid:
                g = make_valid(g)
            return g if not g.is_empty else None
        except Exception:
            try:
                return make_valid(geom)
            except Exception:
                return None

    gdf = gdf.copy()
    gdf["geometry"] = gdf["geometry"].apply(_fix)
    return gdf[gdf["geometry"].notna() & ~gdf["geometry"].is_empty].copy()


def filtrar_poligonos(gdf):
    mask_col = gdf.geometry.geom_type == "GeometryCollection"
    if mask_col.any():
        partes = gdf[mask_col].explode(index_parts=False)
        gdf = gpd.GeoDataFrame(
            pd.concat([gdf[~mask_col], partes], ignore_index=True),
            crs=gdf.crs,
        )
    return gdf[gdf.geometry.geom_type.isin(TIPOS_VALIDOS)].copy()


def atribuir_tile_ids(gdf):
    from shapely.strtree import STRtree

    minx, miny, maxx, maxy = gdf.total_bounds
    tile_defs = bbox_para_tiles(minx, miny, maxx, maxy)

    grid_geoms  = [box(x0, y0, x1, y1) for _, x0, y0, x1, y1 in tile_defs]
    grid_ids    = [t[0] for t in tile_defs]
    grid_cols   = [int(t[0].split("_c")[1]) for t in tile_defs]
    grid_rows   = [int(t[0].split("r")[1].split("_")[0]) for t in tile_defs]

    tree = STRtree(grid_geoms)
    gdf  = gdf.copy()

    tile_id_list  = []
    tile_col_list = []
    tile_row_list = []

    for centroid in gdf.geometry.centroid:
        idx = tree.query(centroid, predicate="within")
        if len(idx) > 0:
            i = idx[0]
        else:
            i = tree.nearest(centroid)
        tile_id_list.append(grid_ids[i])
        tile_col_list.append(grid_cols[i])
        tile_row_list.append(grid_rows[i])

    gdf["tile_id"]  = tile_id_list
    gdf["tile_col"] = tile_col_list
    gdf["tile_row"] = tile_row_list
    return gdf


def processar_camada_vetor(gdf, nome_layer, pasta_out):
    stats = {
        "tipo": "vetor",
        "layer": nome_layer,
        "feicoes_entrada": len(gdf),
        "feicoes_saidas":  0,
        "tiles_gerados":   0,
        "erros":           [],
    }

    # 1. Reproj
    try:
        if gdf.crs is None:
            gdf = gdf.set_crs("EPSG:4326")
        gdf = gdf.to_crs(CRS_ALVO)
    except Exception as e:
        stats["erros"].append("Reproj: " + str(e))
        return stats

    # 2. Filtro
    gdf = filtrar_poligonos(gdf)
    if gdf.empty:
        return stats

    # 3. Repair
    gdf = reparar_geometria(gdf)
    if gdf.empty:
        return stats

    # 4. Tile IDs
    gdf = atribuir_tile_ids(gdf)
    stats["feicoes_saidas"] = len(gdf)

    # 5. Salva por tile
    pasta_out.mkdir(parents=True, exist_ok=True)
    for tile_id, grupo in gdf.groupby("tile_id"):
        if grupo.empty:
            continue
        nome_arquivo = nome_layer + "__" + tile_id + ".parquet"
        try:
            grupo.to_parquet(pasta_out / nome_arquivo, index=False)
            stats["tiles_gerados"] += 1
        except Exception as e:
            stats["erros"].append("Salvar " + nome_arquivo + ": " + str(e))

    return stats


# ---------------------------------------------------------------------------
# RASTERS
# ---------------------------------------------------------------------------

def _resolucao_reproj(src_path):
    """Estima resolucao do raster reprojetado para ESRI:102033."""
    with rasterio.open(src_path) as src:
        transform, width, height = calculate_default_transform(
            src.crs, CRS_RASTERIO, src.width, src.height, *src.bounds
        )
    return transform.a  # tamanho do pixel em metros


def processar_raster(tif_path, nome_base, pasta_out):
    """
    Pipeline raster:
      1. Reproj para ESRI:102033 em memoria (ou arquivo temp se muito grande)
      2. Corta em tiles 25x25 km alinhados ao grid global
      3. Salva cada tile como COG (GeoTIFF com tiles internos COG_TILE_PX x COG_TILE_PX)
         com compressao DEFLATE + predictor=2 (otimo para dados continuos)
    """
    stats = {
        "tipo": "raster",
        "layer": nome_base,
        "feicoes_entrada": 0,
        "feicoes_saidas":  0,
        "tiles_gerados":   0,
        "erros":           [],
    }

    pasta_out.mkdir(parents=True, exist_ok=True)
    tmp_reproj = Path(tempfile.mktemp(suffix="_reproj.tif"))

    try:
        # ---- Passo 1: reprojetar ----
        with rasterio.open(tif_path) as src:
            n_bands = src.count
            dtype   = src.dtypes[0]
            nodata  = src.nodata

            transform_dst, width_dst, height_dst = calculate_default_transform(
                src.crs, CRS_RASTERIO, src.width, src.height, *src.bounds
            )

            profile = src.profile.copy()
            profile.update({
                "crs":       CRS_RASTERIO,
                "transform": transform_dst,
                "width":     width_dst,
                "height":    height_dst,
                "driver":    "GTiff",
                "compress":  "DEFLATE",
                "predictor": 2,
                "tiled":     True,
                "blockxsize": COG_TILE_PX,
                "blockysize": COG_TILE_PX,
                "bigtiff":   "IF_SAFER",
            })
            if nodata is None:
                # Define nodata se ausente
                if np.issubdtype(np.dtype(dtype), np.floating):
                    nodata = float("nan")
                else:
                    nodata = 0
                profile["nodata"] = nodata

            log.info("  Reprojetando: " + tif_path.name)
            with rasterio.open(tmp_reproj, "w", **profile) as dst:
                for band_idx in range(1, n_bands + 1):
                    reproject(
                        source=rasterio.band(src, band_idx),
                        destination=rasterio.band(dst, band_idx),
                        src_transform=src.transform,
                        src_crs=src.crs,
                        dst_transform=transform_dst,
                        dst_crs=CRS_RASTERIO,
                        resampling=Resampling.nearest,
                    )

        # ---- Passo 2: cortar em tiles 25x25 km ----
        with rasterio.open(tmp_reproj) as src_r:
            bounds   = src_r.bounds
            res      = src_r.transform.a  # tamanho pixel em metros
            n_bands  = src_r.count
            dtype    = src_r.dtypes[0]
            nodata   = src_r.nodata
            profile  = src_r.profile.copy()

            log.info(
                "  Resolucao: {:.1f} m  |  Bandas: {}  |  Bounds: {:.0f},{:.0f} -> {:.0f},{:.0f}".format(
                    res, n_bands, bounds.left, bounds.bottom, bounds.right, bounds.top
                )
            )

            tile_defs = bbox_para_tiles(
                bounds.left, bounds.bottom, bounds.right, bounds.top
            )
            log.info("  Tiles a gerar: " + str(len(tile_defs)))

            for tile_id, x0, y0, x1, y1 in tile_defs:
                nome_arquivo = nome_base + "__" + tile_id + ".tif"
                caminho_out  = pasta_out / nome_arquivo

                if caminho_out.exists():
                    stats["tiles_gerados"] += 1
                    continue

                # Janela do tile no espaco do raster reprojetado
                window = rasterio.windows.from_bounds(
                    x0, y0, x1, y1,
                    transform=src_r.transform,
                )
                win_int = window.round_lengths().round_offsets()

                # Verifica se a janela tem intersecao real com o raster
                col_off = int(win_int.col_off)
                row_off = int(win_int.row_off)
                w_width  = int(win_int.width)
                w_height = int(win_int.height)

                col_off  = max(col_off, 0)
                row_off  = max(row_off, 0)
                w_width  = min(w_width,  src_r.width  - col_off)
                w_height = min(w_height, src_r.height - row_off)

                if w_width <= 0 or w_height <= 0:
                    continue  # tile fora da extensao do raster

                win_clipped = rasterio.windows.Window(col_off, row_off, w_width, w_height)
                win_transform = src_r.window_transform(win_clipped)

                # Numero de pixels esperado num tile 25km
                px_por_tile = int(math.ceil(TILE_M / res))

                tile_profile = profile.copy()
                tile_profile.update({
                    "width":      w_width,
                    "height":     w_height,
                    "transform":  win_transform,
                    "driver":     "GTiff",
                    "compress":   "DEFLATE",
                    "predictor":  2,
                    "tiled":      True,
                    "blockxsize": COG_TILE_PX,
                    "blockysize": COG_TILE_PX,
                    "bigtiff":    "IF_SAFER",
                })

                try:
                    data = src_r.read(window=win_clipped)

                    # Pula tiles so com nodata
                    if nodata is not None:
                        if np.isnan(nodata):
                            all_nodata = np.all(np.isnan(data))
                        else:
                            all_nodata = np.all(data == nodata)
                        if all_nodata:
                            continue

                    with rasterio.open(caminho_out, "w", **tile_profile) as dst_t:
                        dst_t.write(data)

                    # Adiciona overviews internas para acelerar zonal stats
                    # em diferentes escalas de zoom
                    _adicionar_overviews(caminho_out, res)

                    stats["tiles_gerados"] += 1

                except Exception as e:
                    stats["erros"].append("Tile " + tile_id + ": " + str(e))

        stats["feicoes_saidas"] = stats["tiles_gerados"]

    except Exception as e:
        stats["erros"].append("Pipeline raster: " + str(e))
        log.error("  ERRO em " + nome_base + ": " + str(e))
    finally:
        if tmp_reproj.exists():
            tmp_reproj.unlink()

    return stats


def _adicionar_overviews(tif_path, resolucao_m):
    """
    Adiciona overviews (piramides) internas no GeoTIFF.
    Fatores calculados para cobrir escalas uteis em zonal stats.
    Com 64 GB de RAM usa DEFLATE + resampling nearest.
    """
    # Fatores de overview: 2, 4, 8, 16 (suficiente para tiles 25km)
    fatores = [2, 4, 8, 16]
    try:
        with rasterio.open(tif_path, "r+") as ds:
            ds.build_overviews(fatores, Resampling.nearest)
            ds.update_tags(ns="rio_overview", resampling="nearest")
    except Exception as e:
        log.warning("  Overviews: " + str(e))


# ---------------------------------------------------------------------------
# Leitura de camadas no ZIP
# ---------------------------------------------------------------------------

def listar_conteudo_zip(zip_path, tmpdir):
    """
    Extrai o ZIP e retorna:
      vetores: lista de {'path', 'driver', 'layer'}
      rasters:  lista de Path para TIFs
    """
    vetores = []
    rasters = []

    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(tmpdir)

    for shp in tmpdir.rglob("*.shp"):
        vetores.append({"path": shp, "driver": "ESRI Shapefile", "layer": None})

    for gpkg in tmpdir.rglob("*.gpkg"):
        try:
            layers = gpd.io.file.fiona.listlayers(str(gpkg))
            for lyr in layers:
                vetores.append({"path": gpkg, "driver": "GPKG", "layer": lyr})
        except Exception:
            vetores.append({"path": gpkg, "driver": "GPKG", "layer": None})

    for ext in ("*.tif", "*.tiff", "*.TIF", "*.TIFF"):
        for tif in tmpdir.rglob(ext):
            rasters.append(tif)

    return vetores, rasters


def ler_camada_vetor(info):
    try:
        kwargs = {"filename": str(info["path"])}
        if info["layer"]:
            kwargs["layer"] = info["layer"]
        gdf = gpd.read_file(**kwargs)
        if gdf.empty or gdf.geometry.isna().all():
            return None
        return gdf
    except Exception as e:
        log.warning("Falha ao ler " + info["path"].name + ": " + str(e))
        return None


# ---------------------------------------------------------------------------
# Processamento de um ZIP
# ---------------------------------------------------------------------------

def processar_zip(zip_path, pasta_vet, pasta_ras):
    all_stats = []
    tmpdir    = Path(tempfile.mkdtemp(prefix="prodes_"))

    try:
        log.info("Extraindo: " + zip_path.name)
        vetores, rasters = listar_conteudo_zip(zip_path, tmpdir)

        if not vetores and not rasters:
            log.warning(zip_path.name + ": nenhum SHP, GPKG ou TIF encontrado.")
            return all_stats

        # --- Vetores ---
        for info in vetores:
            label      = info["layer"] or info["path"].stem
            nome_layer = (zip_path.stem + "__" + label).replace(" ","_").replace("/","_").replace("\\","_")
            log.info("  Vetor: " + nome_layer)
            gdf = ler_camada_vetor(info)
            if gdf is None:
                continue
            s = processar_camada_vetor(gdf, nome_layer, pasta_vet)
            all_stats.append(s)
            log.info(
                "    " + str(s["feicoes_entrada"]) + " feat -> "
                + str(s["feicoes_saidas"]) + " validas -> "
                + str(s["tiles_gerados"]) + " tiles"
            )

        # --- Rasters ---
        for tif_path in rasters:
            nome_base = (zip_path.stem + "__" + tif_path.stem).replace(" ","_").replace("/","_").replace("\\","_")
            log.info("  Raster: " + nome_base)
            s = processar_raster(tif_path, nome_base, pasta_ras)
            all_stats.append(s)
            log.info(
                "    " + str(s["tiles_gerados"]) + " tiles gerados"
                + (" | ERROS: " + str(len(s["erros"])) if s["erros"] else "")
            )

    except zipfile.BadZipFile:
        log.error(zip_path.name + ": ZIP corrompido, pulando.")
    except Exception as e:
        log.error(zip_path.name + ": erro inesperado: " + str(e))
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    return all_stats


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    zips = sorted(PASTA_ZIPS.glob("*.zip"))

    if not zips:
        log.error("Nenhum ZIP encontrado em: " + str(PASTA_ZIPS))
        sys.exit(1)

    log.info("Pasta entrada  : " + str(PASTA_ZIPS))
    log.info("Saida vetores  : " + str(PASTA_VET))
    log.info("Saida rasters  : " + str(PASTA_RAS))
    log.info("ZIPs           : " + str(len(zips)))
    log.info("CRS alvo       : " + CRS_ALVO)
    log.info("Tile           : " + str(TILE_M // 1000) + " km")
    log.info("COG tile px    : " + str(COG_TILE_PX) + "x" + str(COG_TILE_PX))
    log.info("Threads        : " + str(MAX_WORKERS))
    print()

    PASTA_VET.mkdir(parents=True, exist_ok=True)
    PASTA_RAS.mkdir(parents=True, exist_ok=True)

    todos_stats = []

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futuros = {
            executor.submit(processar_zip, z, PASTA_VET, PASTA_RAS): z
            for z in zips
        }
        with tqdm(total=len(zips), unit="zip", ncols=80, desc="ZIPs") as bar:
            for fut in as_completed(futuros):
                zip_path = futuros[fut]
                try:
                    todos_stats.extend(fut.result())
                except Exception as e:
                    log.error(zip_path.name + ": " + str(e))
                finally:
                    bar.update(1)

    # Resumo
    vet_stats = [s for s in todos_stats if s["tipo"] == "vetor"]
    ras_stats = [s for s in todos_stats if s["tipo"] == "raster"]

    print()
    sep = "=" * 60
    log.info(sep)
    log.info("RESUMO FINAL")
    log.info(sep)
    log.info("ZIPs processados    : " + str(len(zips)))
    log.info("--- Vetores ---")
    log.info("  Camadas           : " + str(len(vet_stats)))
    log.info("  Features salvas   : " + str(sum(s["feicoes_saidas"] for s in vet_stats)))
    log.info("  Tiles GeoParquet  : " + str(sum(s["tiles_gerados"]  for s in vet_stats)))
    log.info("--- Rasters ---")
    log.info("  Arquivos TIF      : " + str(len(ras_stats)))
    log.info("  Tiles COG         : " + str(sum(s["tiles_gerados"]  for s in ras_stats)))
    log.info("  Erros totais      : " + str(sum(len(s["erros"])     for s in todos_stats)))
    log.info("Saida vetores       : " + str(PASTA_VET))
    log.info("Saida rasters       : " + str(PASTA_RAS))
    log.info(sep)

    # Relatorio JSON
    relatorio_path = PASTA_ZIPS / "relatorio_processamento.json"
    with open(relatorio_path, "w", encoding="utf-8") as f:
        json.dump(
            {
                "processado_em": datetime.utcnow().isoformat(),
                "pasta_entrada": str(PASTA_ZIPS),
                "pasta_vetores": str(PASTA_VET),
                "pasta_rasters": str(PASTA_RAS),
                "crs":           CRS_ALVO,
                "tile_km":       TILE_M / 1000,
                "cog_tile_px":   COG_TILE_PX,
                "camadas":       todos_stats,
            },
            f, ensure_ascii=False, indent=2, default=str,
        )
    log.info("Relatorio salvo: " + str(relatorio_path))


if __name__ == "__main__":
    main()