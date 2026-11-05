# MAC-Fabricante

Automação em Python para identificar fabricantes de equipamentos em uma planilha Excel a partir do endereço MAC. O script normaliza os endereços, extrai o prefixo OUI e consulta o serviço MAC Vendors quando o prefixo ainda não está no cache local.

## Experimentar com dados fictícios

Na raiz do projeto, com Python 3 e acesso à internet:

```powershell
python -m pip install pandas openpyxl requests
python MAC_Fabricante.py
```

O script lê `Login_inicio.xlsx` e grava `login_final.xlsx` com as colunas `Login`, `MAC` e `Fabricante`. As planilhas versionadas são exemplos fictícios. Rodar o script pode sobrescrever o exemplo de saída.

## Usar outra planilha

Prepare um XLSX com `Login` e `MAC` e ajuste os nomes de entrada e saída na chamada a `main` no final de `MAC_Fabricante.py`. O arquivo local `oui_cache.json` evita novas consultas para prefixos já conhecidos. Mantenha exportações reais do provedor fora do repositório.
