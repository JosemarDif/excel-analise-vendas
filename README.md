# 📗 Análise de Vendas Angola — Excel Avançado

![Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=flat&logo=microsoftexcel&logoColor=white)
![Status](https://img.shields.io/badge/Status-Concluído-4ADE80?style=flat)
![Registos](https://img.shields.io/badge/Dados-100%20Transações-217346?style=flat)

> Ficheiro Excel completo com 100 transações de vendas, fórmulas RANK e IF, formatação condicional por meta, gráficos profissionais de barras e pizza por vendedor e produto.

---

## 🎯 Objetivos

- Registar e organizar 100 transações de vendas
- Calcular automaticamente se cada venda atingiu a meta
- Fazer o ranking dos vendedores por performance
- Visualizar a distribuição de vendas com gráficos

---

## 📁 Estrutura do Ficheiro

```
Analise_Vendas_Angola.xlsx
├── 📊 Dados Vendas         ← 100 transações completas
├── 👤 Resumo Vendedor      ← KPIs por vendedor + gráfico barras
└── 📦 Resumo por Produto   ← Receita por produto + gráfico pizza
```

---

## 🔧 Fórmulas Utilizadas

```excel
-- Calcular Total da Venda
=H2*I2

-- Verificar se Meta foi Atingida
=IF(J2>=K2,"Sim","Não")

-- Ranking do Vendedor (sem empates)
=RANK(B4,B$4:B$8,0)

-- Percentagem do Total
=B4/B$9

-- Ticket Médio
=B4/C4

-- Total Geral
=SUM(B4:B8)
```

---

## 🎨 Formatação Condicional

| Condição | Cor Aplicada | Coluna |
|----------|-------------|--------|
| Meta atingida ("Sim") | 🟢 Fundo verde `#C5E0B4` | Meta Atingida |
| Meta não atingida ("Não") | 🔴 Fundo vermelho `#FFD7D7` | Meta Atingida |
| Top 1 em ranking | ⭐ Dourado `#F5A623` | Total Vendas |

---

## 📊 KPIs Principais

| KPI | Valor |
|-----|-------|
| Total de Vendas | 21.1M Kz |
| Melhor Vendedor | Maria João (5.1M Kz) |
| Produto mais vendido | Laptop (8.4M Kz) |
| Ticket Médio Geral | 180.341 Kz |

---

## 🏷️ Tecnologias

`Excel` `IF` `RANK` `SUM` `AVERAGE` `Formatação Condicional` `Gráfico de Barras` `Gráfico de Pizza` `Tabelas Estilizadas`

---

*Autor: **Josemar Manuel** · [LinkedIn](https://linkedin.com/in/josemarmanuel) · Luanda, Angola*
