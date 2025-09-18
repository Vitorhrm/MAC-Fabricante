## 1. Origem do Relatório (IXC Provedor)
Caminho no sistema:
```
Provedor -> Logins
```
Passos:
1. Marque o filtro para exibir **somente logins ativos**.
2. Gere / exporte o relatório em formato **XLSX**.
3. Abra a planilha exportada.
4. Mantenha apenas a coluna de identificação do usuário (ex: `Login` ou `IP`) e a coluna de MAC.

5. Salve o arquivo XLSX.

Requisitos mínimos da planilha para o script:
- Coluna `MAC` (obrigatória)
- Coluna `Login` (para sair no resultado final)

## 2. O que o Script Faz
1. Lê o arquivo de entrada Excel.
2. Normaliza os MACs (remove separadores e deixa apenas 12 hex em maiúsculas).
3. Extrai o OUI (primeiros 6 hex) de cada MAC.
4. Usa cache local (`oui_cache.json`) para evitar consultas repetidas.
5. Consulta a API do MAC Vendors apenas para OUIs novos.
6. Atribui o fabricante (`Fabricante`).
7. Ordena pelos valores numéricos dos MACs.
8. Gera um novo Excel contendo apenas as colunas: `Login`, `MAC`, `Fabricante`.

## 3. Arquivos Principais
- `MAC_Fabricante.py`: Script principal.
- `oui_cache.json`: Cache persistente (criado após a primeira execução).
- Entrada e saida pode mudar conforme parâmetro.

## 4. Pré-requisitos
- Python 3.11+ (recomendado)
- Pacotes:
  - pandas
  - requests

Instalação rápida (PowerShell):
```powershell
pip install pandas requests
```

## 5. Execução Básica
Assumindo que o arquivo de entrada se chama `Login_inicio.xlsx`:
```powershell
python MAC_Fabricante.py
```
Saída gerada: `login_final.xlsx`.

Se quiser mudar nomes, edite a chamada no final do script:
```python
if __name__ == "__main__":
    main("entrada.xlsx", "saida.xlsx")
```

## 6. Cache de OUIs
- Local: `oui_cache.json`
- Estrutura: dicionário JSON chave = OUI (6 hex), valor = nome do fabricante.
- Benefício: evita repetir chamadas (economiza tempo e limite de taxa).

Se quiser começar do zero, basta remover o arquivo de cache:

## 7. Tratamento de Erros / Casos Especiais
- MAC inválido ou muito curto: linha resultará em `Fabricante` vazio (ou mapeado como `Desconhecido`).
- API retorna 404 / 204: fabricante definido como `Desconhecido`.
- Rate limit (429): script aguarda e tenta novamente.


## 9. Estrutura do Resultado
Colunas finais:
```
Login | MAC | Fabricante
```

## 10. Boas Práticas no Preparo do XLSX
- Remover colunas desnecessárias (reduz memória e acelera).
- Garantir que não há espaços extras no cabeçalho (`MAC`, não ` MAC`).
- Verificar se não existem MACs duplicados inesperados.
