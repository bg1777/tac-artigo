import zipfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
BRUTOS = RAIZ / "dados_brutos"
EXTRAIDOS = RAIZ / "dados_extraidos"

for z in sorted(BRUTOS.glob("*.zip")):
    destino = EXTRAIDOS / z.stem
    if destino.exists() and any(destino.iterdir()):
        print(f"[pulando] {z.name} (já extraído)")
        continue
    print(f"[extraindo] {z.name}")
    with zipfile.ZipFile(z) as f:
        f.extractall(destino)