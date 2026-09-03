# Release e distribuição
A versão está em core/version.py, VERSION, pyproject.toml e no script do instalador.
Build local:
1. cd central_n2
2. python -m pip install -e ".[dev]"
3. python -m pytest -q
4. pyinstaller CentralN2.spec
Instalador: use Inno Setup com installer/CentralN2.iss.
Release automatizado: após CI verde, crie uma tag v5.0.0 e envie ao GitHub. O workflow testa antes de publicar artefatos.
Configuração por ambiente: não embuta settings.local.json no release. Crie-o no ambiente de destino a partir do exemplo.
