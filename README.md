# Relatório Escolas por Município do Ceará

Relatório interativo em HTML com mapa municipal do Ceará, filtros e dashboards de quantidade de escolas e alunos.

Também inclui abas baseadas em `Mapeamento`, reutilizando a mesma malha municipal para visualizar distribuição territorial, regiões, batalhões, funções, postos/graduações e a relação de antiguidade com localização no mapa.

## Rodar localmente

```bash
python3 main.py
```

Acesse:

```text
http://127.0.0.1:5010
```

## Gerar versão para GitHub Pages

Sempre que a planilha `Análise.xlsx` mudar, gere novamente os arquivos estáticos localmente:

```bash
python3 build_static.py
```

Isso atualiza:

- `docs/index.html`
- `docs/ceara_municipios.pdf`
- `docs/ceara_municipios.csv`

No GitHub, o workflow `.github/workflows/update-static-site.yml` roda esse build automaticamente quando houver push na `main` alterando template, dados ou scripts do relatório. Se `docs/` mudar, ele cria um commit automático com a versão estática atualizada.

## Publicar online no GitHub Pages

1. Envie este projeto para um repositório no GitHub.
2. No GitHub, abra `Settings` > `Pages`.
3. Em `Build and deployment`, escolha `Deploy from a branch`.
4. Selecione a branch `main` e a pasta `/docs`.
5. Salve.

O GitHub vai gerar um link parecido com:

```text
https://seu-usuario.github.io/nome-do-repositorio/
```
