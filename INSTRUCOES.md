# 🏷️ Gerador de Dinâmicas — Guia Completo

## Fluxo semanal (2 minutos)
1. Abrir o link do app
2. Subir o background da semana
3. Subir o Excel da lâmina
4. Clicar em Gerar → baixar ZIP ou PDF

## Links fixos no Drive (configura uma vez)
- Gagalin-Regular.otf (fonte do nome)
- chunkfive.regular.ttf (fonte do preço)
- Box de preço PNG (opcional)
- Pasta de fotos dos produtos

## Regras automáticas do Excel
- Coluna B = nome | Coluna H = preço
- Nomes com SELL IN/OUT, PDN, CUSTO FINAL etc. são limpos automaticamente
- KG: peça, pedaço, fatiada, a granel, presunto cozido, costela suína, queijo de coalho godam
- CADA: todo o resto
- Fórmulas (=2.99*18) são ignoradas

## Layout da dinâmica (A4 300dpi)
- Background: ocupa o topo (~55%)
- Nome do produto: 11,5cm do topo, preto, centralizado, até 2 linhas
- Foto do produto: centralizada abaixo do nome
- Box de preço: centro em x=15,5cm, y=24,5cm (canto inferior direito)
- Dentro do box: R$ pequeno (topo esq) | valor grande centralizado | unidade abaixo à direita

## Nomear imagens no Drive
Nome = slug do produto: minúsculas, sem acento, espaços viram _
Ex: "Bife Ancho Bovino Maturatta Peça Kg" → bife_ancho_bovino_maturatta_peca_kg.png

## Publicar no Streamlit Cloud (grátis)
1. github.com → criar repositório
2. Subir app.py e requirements.txt
3. share.streamlit.io → conectar GitHub → Deploy
4. Compartilhar link com a equipe

## Compartilhar arquivos do Drive
Botão direito → Compartilhar → "Qualquer pessoa com o link" → Leitor → Copiar link
